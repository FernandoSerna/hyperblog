VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_comportamiento_hora_hora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Piezas leidas hora - hora          F5 Para ver piezas leidas por usuario   F7 Para ver grafica de aprendizaje"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   3210
      Left            =   30
      TabIndex        =   12
      Top             =   7215
      Width           =   12180
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   12285
         TabIndex        =   13
         Top             =   525
         Width           =   1650
      End
      Begin MSComctlLib.ListView lv_grafica_3 
         Height          =   3000
         Left            =   105
         TabIndex        =   14
         Top             =   135
         Width           =   11940
         _ExtentX        =   21061
         _ExtentY        =   5292
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "   Usuario"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "15 - 16"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "16 - 17"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "17 - 18"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "18 - 19"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "19 - 20"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "20 - 21"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "21 - 22"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "22 - 23"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "TOTAL"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "USUARIO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3210
      Left            =   30
      TabIndex        =   3
      Top             =   3990
      Width           =   12180
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   12240
         TabIndex        =   4
         Top             =   525
         Width           =   1650
      End
      Begin MSComctlLib.ListView lv_grafica_2 
         Height          =   3000
         Left            =   120
         TabIndex        =   5
         Top             =   135
         Width           =   11940
         _ExtentX        =   21061
         _ExtentY        =   5292
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "   Usuario"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "7 - 8"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "8 - 9"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "9 - 10"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "10 - 11"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "11 - 12"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "12 - 13"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "13 - 14"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "14 - 15"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "TOTAL"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "USUARIO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   240
   End
   Begin VB.Frame Frame2 
      Height          =   3210
      Left            =   15
      TabIndex        =   0
      Top             =   720
      Width           =   12180
      Begin VB.TextBox txt_foco 
         Height          =   315
         Left            =   12315
         TabIndex        =   1
         Top             =   525
         Width           =   1650
      End
      Begin MSComctlLib.ListView lv_grafica 
         Height          =   3000
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   11940
         _ExtentX        =   21061
         _ExtentY        =   5292
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "   Usuario"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "23 - 24"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "0 - 1"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "1 - 2"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "2 - 3"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "3 - 4 "
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "4 - 5"
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "5 - 6 "
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "6 - 7 "
            Object.Width           =   1632
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "TOTAL"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "USUARIO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   30
      TabIndex        =   6
      Top             =   -45
      Width           =   12165
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   6480
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmd_pausa 
         Height          =   315
         Left            =   4455
         Picture         =   "frmoracle_comportamiento_hora_hora.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmd_play 
         Height          =   315
         Left            =   4140
         Picture         =   "frmoracle_comportamiento_hora_hora.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmd_excel 
         Height          =   315
         Left            =   3810
         Picture         =   "frmoracle_comportamiento_hora_hora.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   315
      End
      Begin VB.TextBox txt_fecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1230
         TabIndex        =   8
         Text            =   "11/10/2012"
         Top             =   135
         Width           =   2040
      End
      Begin VB.Label lbl_total 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9240
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8280
         TabIndex        =   15
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         TabIndex        =   7
         Top             =   225
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmoracle_comportamiento_hora_hora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter


Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_excel_Click()
   x = 0
   If x = 1 Then
        Set reporte = appl.OpenReport(App.Path + "\rep_oracle_comportamiento_hora_hora.rpt")
        reporte.RecordSelectionFormula = "{VW_ORACLE_COMPORTAMIENTO_HORA_HORA.FECHA} = '" + Me.txt_fecha + "' and {VW_ORACLE_COMPORTAMIENTO_HORA_HORA.VCHA_UOR_UNIDAD_ID} = '" + var_unidad_organizacional + "'"
        For ntablas = 1 To reporte.Database.Tables.Count
            reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
        Next ntablas
        reporte.ExportOptions.FormatType = crEFTExcel80
        reporte.ExportOptions.DestinationType = crEDTDiskFile
        archivo = "c:\reportessid\comportamiento_hora_hora_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
        reporte.ExportOptions.DiskFileName = archivo
        reporte.Export False
        Set reporte = Nothing
        MsgBox "Se a terminado de guardar el archivo " + archivo
    Else
        var_dia = CStr(Day(CDate(Me.txt_fecha)))
        var_mes = CStr(Month(CDate(Me.txt_fecha)))
        var_año = CStr(Year(CDate(Me.txt_fecha)))
        If Len(Trim(var_dia)) = 1 Then
           var_dia = "0" + var_dia
        End If
        If Len(Trim(var_mes)) = 1 Then
           var_mes = "0" + var_mes
        End If
        var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
        var_cadena = "SELECT FECHA, VCHA_USU_NOMBRE + ' ' + VCHA_USU_APELLIDOS AS NOMBRE_USUARIO, (SELECT MAX(B.H_23_24) H_23_24 FROM VW_ORACLE_COMPORTAMIENTO_HORA_HORA B WHERE B.USUARIO = A.USUARIO AND CAST(B.FECHA AS DATETIME) = CAST((CAST(A.FECHA AS DATETIME) - 1) AS VARCHAR(50))) H_23_24, H_0_1, H_1_2, H_2_3, H_3_4, H_4_5, H_5_6, H_6_7, H_7_8, H_8_9, H_9_10, H_10_11, H_11_12, H_12_13 , H_13_14, H_14_15, H_15_16, H_16_17, H_17_18, H_18_19, H_19_20, H_20_21, H_21_22, H_22_23  FROM VW_ORACLE_COMPORTAMIENTO_HORA_HORA A WHERE (FECHA = " + var_fecha + ") AND (VCHA_UOR_UNIDAD_ID = " + var_unidad_organizacional + ") ORDER BY H_0_1 DESC, H_1_2 DESC, H_2_3 DESC, H_3_4 DESC, H_4_5 DESC, H_5_6 DESC, H_6_7 DESC, H_7_8 DESC, H_8_9 DESC, H_9_10 DESC, H_10_11 DESC, H_11_12 DESC, H_12_13 DESC, H_13_14 DESC, H_14_15 DESC, H_15_16 DESC, H_16_17 DESC, H_17_18 DESC, H_18_19 DESC, H_19_20 DESC, H_20_21 DESC, H_21_22 DESC, H_22_23 DESC, H_23_24 DESC"
        Dim iFila As Long, iCol As Integer, i As Integer
        Set oexcel = CreateObject("Excel.Application")
        Set owbook = oexcel.Workbooks.Add
        Set osheet = owbook.Worksheets(1)
        osheet.Name = "ASIENTO"
        Screen.MousePointer = vbHourglass
        iFila = 1
        ifila2 = 1
        icol2 = 1
        iCol = 1
        'MsgBox var_cadena
        Text3 = var_cadena
        rsaux10.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
        For i = 0 To rsaux10.Fields.Count - 1
            osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
        Next
        iFila = iFila + 1
        With osheet
             ' carga los registros del recordset
             .Cells(iFila, iCol).CopyFromRecordset rsaux10
             'oExcel.Columns(1).Select
             'oExcel.Selection.NumberFormat = "#,##0"
             oexcel.Columns(1).Select
             oexcel.Selection.Font.Color = vbRed
             .Columns.AutoFit ' ajusta el ancho de las columnas
        End With
        owbook.SaveAs "c:\reportessid\comportamiento_hora_hora_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
        oexcel.Visible = True
        Set oexcel = Nothing
        Screen.MousePointer = vbDefault
        rsaux10.Close
    End If
End Sub

Private Sub cmd_pausa_Click()
   Me.Timer1.Enabled = False
End Sub

Private Sub cmd_play_Click()
   Me.Timer1.Enabled = True
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   Me.txt_fecha = Date
   
   Me.Timer1.Enabled = True
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub lv_grafica_2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_grafica_2, ColumnHeader)
End Sub

Private Sub lv_grafica_2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 118 Then
      If Me.lv_grafica_2.ListItems.Count > 0 Then
         var_usuario_h_x_h = Me.lv_grafica_2.selectedItem.SubItems(10)
         frmoracle_periodo_grafica_aprendizaje.Show 1
      End If
   End If
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         If Me.lv_grafica_2.ListItems.Count > 0 Then
            var_fecha_inicio = CDate(Me.txt_fecha) - 1
            var_año_str = CStr(Year(var_fecha_inicio))
            var_dia_str = CStr(Day(var_fecha_inicio))
            var_mes_str = CStr(Month(var_fecha_inicio))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha_inicio_str = CStr(var_año_str) + "/" + CStr(var_mes_str) + "/" + CStr(var_dia_str) + " 23:00:00"
      
            var_año_str = CStr(Year(CDate(Me.txt_fecha)))
            var_dia_str = CStr(Day(CDate(Me.txt_fecha)))
            var_mes_str = CStr(Month(CDate(Me.txt_fecha)))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha = "{d '" + CStr(var_año_str) + "-" + CStr(var_mes_str) + "-" + CStr(var_dia_str) + "'}"
      
            var_fecha_fin_str = CStr(var_año_str) + "/" + CStr(var_mes_str) + "/" + CStr(var_dia_str) + " 23:00:00"
      
            rs.Open "ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY/MM/DD HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = " select a.codigo, b.descripcion,linea, sum(cantidad) as cantidad from xxvia_tb_bitacora_lectura a, xxvia_vw_articulos_cat b where usuario = ? and fecha_hora >= to_date(?,'YYYY/MM/DD HH24:MI:SS') and fecha_hora < to_date(?,'YYYY/MM/DD HH24:MI:SS') and a.codigo = b.codigo and b.organization_id = 93 group by a.codigo, b.descripcion,linea"
            'var_cadena = "select a.codigo, b.descripcion,linea, To_Number(To_Char(FECHA_HORA, 'HH24')) hora ,sum(cantidad) from xxvia_tb_bitacora_lectura a, xxvia_vw_articulos_cat b where usuario = ? and fecha_hora >= to_date(?,'YYYY/MM/DD HH24:MI:SS') and fecha_hora < to_date(?,'YYYY/MM/DD HH24:MI:SS') and a.codigo = b.codigo and b.organization_id = 93 group by a.codigo, b.descripcion,linea, To_Number(To_Char(FECHA_HORA, 'HH24')) ORDER BY To_Number(To_Char(FECHA_HORA, 'HH24')), Linea"
            strconsulta = var_cadena
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.lv_grafica_2.selectedItem.SubItems(10))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_fecha_inicio_str)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_fecha_fin_str)
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rs.EOF Then
               cnn.BeginTrans
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA ", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux1.Close
               rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_cadena = "INSERT INTO TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA (INTE_tEM_CONSECUTIVO, USUARIO, NOMBRE_USUARIO, CODIGO, DESCRIPCION, LINEA, CANTIDAD, FECHA) VALUES (" + CStr(var_consecutivo) + ",'" + Me.lv_grafica_2.selectedItem.SubItems(10) + "','" + Me.lv_grafica_2.selectedItem + "','" + rs!codigo + "','" + rs!descripcion + "','" + rs!Linea + "'," + CStr(rs!cantidad) + "," + var_fecha + ")"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and codigo is null", cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_lectura_hora_hora.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_LECTURA_USUARIOS_HORA_HORA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte lectura " + Me.lv_grafica_2.selectedItem + " a la fecha " + Me.txt_fecha
               frmvistasprevias.Show 1
               Set reporte = Nothing
         
               rsaux1.Open "SELECT FECHA, USUARIO, NOMBRE_USUARIO, CODIGO, DESCRIPCION, LINEA, CANTIDAD from TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " ORDER BY LINEA", cnn, adOpenDynamic, adLockOptimistic
               Dim iFila As Long, iCol As Integer, i As Integer
               Set oexcel = CreateObject("Excel.Application")
               Set owbook = oexcel.Workbooks.Add
               Set osheet = owbook.Worksheets(1)
               osheet.Name = "COMPORTAMIENTO"
               Screen.MousePointer = vbHourglass
               iFila = 1
               ifila2 = 1
               icol2 = 1
               iCol = 1
               'MsgBox var_cadena
               For i = 0 To rsaux1.Fields.Count - 1
                   osheet.Cells(iFila, i + 1) = rsaux1.Fields(i).Name
                   osheet.Cells(iFila, i + 1).Font.Bold = True
               Next
               iFila = iFila + 1
                
               With osheet
                    ' carga los registros del recordset
                    .Cells(iFila, iCol).CopyFromRecordset rsaux1
                    'oExcel.Columns(1).Select
                    'oExcel.Selection.NumberFormat = "#,##0"
                    'oExcel.Columns(1).Select
                    'oExcel.Selection.Font.Color = vbRed
                    .Columns.AutoFit ' ajusta el ancho de las columnas
               End With
               owbook.SaveAs "c:\reportessid\COMPORTAMIENTO_" + Replace(Me.lv_grafica_2.selectedItem, " ", "_") + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               oexcel.Visible = True
               Set oexcel = Nothing
               Screen.MousePointer = vbDefault
               rsaux1.Close
               rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         
            Else
               MsgBox "No existe resultados para el usuario seleccionado", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "No se a seleccionado un usuario", vbOKOnly, "ATENCION"
         End If
      '
      Else
         MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub lv_grafica_3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_grafica_3, ColumnHeader)
End Sub

Private Sub lv_grafica_3_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 118 Then
      If Me.lv_grafica_3.ListItems.Count > 0 Then
         var_usuario_h_x_h = Me.lv_grafica_3.selectedItem.SubItems(10)
         frmoracle_periodo_grafica_aprendizaje.Show 1
      End If
   End If
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         If Me.lv_grafica_3.ListItems.Count > 0 Then
            var_fecha_inicio = CDate(Me.txt_fecha) - 1
            var_año_str = CStr(Year(var_fecha_inicio))
            var_dia_str = CStr(Day(var_fecha_inicio))
            var_mes_str = CStr(Month(var_fecha_inicio))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha_inicio_str = CStr(var_año_str) + "/" + CStr(var_mes_str) + "/" + CStr(var_dia_str) + " 23:00:00"
      
            var_año_str = CStr(Year(CDate(Me.txt_fecha)))
            var_dia_str = CStr(Day(CDate(Me.txt_fecha)))
            var_mes_str = CStr(Month(CDate(Me.txt_fecha)))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha = "{d '" + CStr(var_año_str) + "-" + CStr(var_mes_str) + "-" + CStr(var_dia_str) + "'}"
      
            var_fecha_fin_str = CStr(var_año_str) + "/" + CStr(var_mes_str) + "/" + CStr(var_dia_str) + " 23:00:00"
      
            rs.Open "ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY/MM/DD HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = " select a.codigo, b.descripcion,linea, sum(cantidad) as cantidad from xxvia_tb_bitacora_lectura a, xxvia_vw_articulos_cat b where usuario = ? and fecha_hora >= to_date(?,'YYYY/MM/DD HH24:MI:SS') and fecha_hora < to_date(?,'YYYY/MM/DD HH24:MI:SS') and a.codigo = b.codigo and b.organization_id = 93 group by a.codigo, b.descripcion,linea"
            'var_cadena = "select a.codigo, b.descripcion,linea, To_Number(To_Char(FECHA_HORA, 'HH24')) hora ,sum(cantidad) from xxvia_tb_bitacora_lectura a, xxvia_vw_articulos_cat b where usuario = ? and fecha_hora >= to_date(?,'YYYY/MM/DD HH24:MI:SS') and fecha_hora < to_date(?,'YYYY/MM/DD HH24:MI:SS') and a.codigo = b.codigo and b.organization_id = 93 group by a.codigo, b.descripcion,linea, To_Number(To_Char(FECHA_HORA, 'HH24')) ORDER BY To_Number(To_Char(FECHA_HORA, 'HH24')), Linea"
            strconsulta = var_cadena
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.lv_grafica_3.selectedItem.SubItems(10))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_fecha_inicio_str)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_fecha_fin_str)
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rs.EOF Then
               cnn.BeginTrans
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA ", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux1.Close
               rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_cadena = "INSERT INTO TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA (INTE_tEM_CONSECUTIVO, USUARIO, NOMBRE_USUARIO, CODIGO, DESCRIPCION, LINEA, CANTIDAD, FECHA) VALUES (" + CStr(var_consecutivo) + ",'" + Me.lv_grafica_3.selectedItem.SubItems(10) + "','" + Me.lv_grafica_3.selectedItem + "','" + rs!codigo + "','" + rs!descripcion + "','" + rs!Linea + "'," + CStr(rs!cantidad) + "," + var_fecha + ")"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and codigo is null", cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_lectura_hora_hora.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_LECTURA_USUARIOS_HORA_HORA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte lectura " + Me.lv_grafica_3.selectedItem + " a la fecha " + Me.txt_fecha
               frmvistasprevias.Show 1
               Set reporte = Nothing
         
               rsaux1.Open "SELECT FECHA, USUARIO, NOMBRE_USUARIO, CODIGO, DESCRIPCION, LINEA, CANTIDAD from TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " ORDER BY LINEA", cnn, adOpenDynamic, adLockOptimistic
               Dim iFila As Long, iCol As Integer, i As Integer
               Set oexcel = CreateObject("Excel.Application")
               Set owbook = oexcel.Workbooks.Add
               Set osheet = owbook.Worksheets(1)
               osheet.Name = "COMPORTAMIENTO"
               Screen.MousePointer = vbHourglass
               iFila = 1
               ifila2 = 1
               icol2 = 1
               iCol = 1
               'MsgBox var_cadena
               For i = 0 To rsaux1.Fields.Count - 1
                   osheet.Cells(iFila, i + 1) = rsaux1.Fields(i).Name
                   osheet.Cells(iFila, i + 1).Font.Bold = True
               Next
               iFila = iFila + 1
                
               With osheet
                    ' carga los registros del recordset
                    .Cells(iFila, iCol).CopyFromRecordset rsaux1
                    'oExcel.Columns(1).Select
                    'oExcel.Selection.NumberFormat = "#,##0"
                    'oExcel.Columns(1).Select
                    'oExcel.Selection.Font.Color = vbRed
                    .Columns.AutoFit ' ajusta el ancho de las columnas
               End With
               owbook.SaveAs "c:\reportessid\COMPORTAMIENTO_" + Replace(Me.lv_grafica_3.selectedItem, " ", "_") + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               oexcel.Visible = True
               Set oexcel = Nothing
               Screen.MousePointer = vbDefault
               rsaux1.Close
               rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         
            Else
               MsgBox "No existe resultados para el usuario seleccionado", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "No se a seleccionado un usuario", vbOKOnly, "ATENCION"
         End If
      '
      Else
         MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
      End If
   End If

End Sub

Private Sub lv_grafica_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_grafica, ColumnHeader)
End Sub

Private Sub lv_grafica_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         If Me.lv_grafica.ListItems.Count > 0 Then
            var_fecha_inicio = CDate(Me.txt_fecha) - 1
            var_año_str = CStr(Year(var_fecha_inicio))
            var_dia_str = CStr(Day(var_fecha_inicio))
            var_mes_str = CStr(Month(var_fecha_inicio))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha_inicio_str = CStr(var_año_str) + "/" + CStr(var_mes_str) + "/" + CStr(var_dia_str) + " 23:00:00"
      
            var_año_str = CStr(Year(CDate(Me.txt_fecha)))
            var_dia_str = CStr(Day(CDate(Me.txt_fecha)))
            var_mes_str = CStr(Month(CDate(Me.txt_fecha)))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha = "{d '" + CStr(var_año_str) + "-" + CStr(var_mes_str) + "-" + CStr(var_dia_str) + "'}"
      
            var_fecha_fin_str = CStr(var_año_str) + "/" + CStr(var_mes_str) + "/" + CStr(var_dia_str) + " 23:00:00"
      
            rs.Open "ALTER SESSION SET NLS_DATE_FORMAT = 'YYYY/MM/DD HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = " select a.codigo, b.descripcion,linea, sum(cantidad) as cantidad from xxvia_tb_bitacora_lectura a, xxvia_vw_articulos_cat b where usuario = ? and fecha_hora >= to_date(?,'YYYY/MM/DD HH24:MI:SS') and fecha_hora < to_date(?,'YYYY/MM/DD HH24:MI:SS') and a.codigo = b.codigo and b.organization_id = 93 group by a.codigo, b.descripcion,linea"
            'var_cadena = "select a.codigo, b.descripcion,linea, To_Number(To_Char(FECHA_HORA, 'HH24')) hora ,sum(cantidad) from xxvia_tb_bitacora_lectura a, xxvia_vw_articulos_cat b where usuario = ? and fecha_hora >= to_date(?,'YYYY/MM/DD HH24:MI:SS') and fecha_hora < to_date(?,'YYYY/MM/DD HH24:MI:SS') and a.codigo = b.codigo and b.organization_id = 93 group by a.codigo, b.descripcion,linea, To_Number(To_Char(FECHA_HORA, 'HH24')) ORDER BY To_Number(To_Char(FECHA_HORA, 'HH24')), Linea"
            strconsulta = var_cadena
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Me.lv_grafica.selectedItem.SubItems(10))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_fecha_inicio_str)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_fecha_fin_str)
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rs.EOF Then
               cnn.BeginTrans
               If rsaux1.State = 1 Then
                  rsaux1.Close
               End If
               rsaux1.Open "SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA ", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux1.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux1.Close
               rsaux1.Open "INSERT INTO TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rs.EOF
                     var_cadena = "INSERT INTO TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA (INTE_tEM_CONSECUTIVO, USUARIO, NOMBRE_USUARIO, CODIGO, DESCRIPCION, LINEA, CANTIDAD, FECHA) VALUES (" + CStr(var_consecutivo) + ",'" + Me.lv_grafica.selectedItem.SubItems(10) + "','" + Me.lv_grafica.selectedItem + "','" + rs!codigo + "','" + rs!descripcion + "','" + rs!Linea + "'," + CStr(rs!cantidad) + "," + var_fecha + ")"
                     rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rsaux1.Open "delete from TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " and codigo is null", cnn, adOpenDynamic, adLockOptimistic
               Set reporte = appl.OpenReport(App.Path + "\rep_oracle_lectura_hora_hora.rpt")
               reporte.RecordSelectionFormula = "{VW_ORACLE_LECTURA_USUARIOS_HORA_HORA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte lectura " + Me.lv_grafica.selectedItem + " a la fecha " + Me.txt_fecha
               frmvistasprevias.Show 1
               Set reporte = Nothing
         
               rsaux1.Open "SELECT FECHA, USUARIO, NOMBRE_USUARIO, CODIGO, DESCRIPCION, LINEA, CANTIDAD from TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo) + " ORDER BY LINEA", cnn, adOpenDynamic, adLockOptimistic
               Dim iFila As Long, iCol As Integer, i As Integer
               Set oexcel = CreateObject("Excel.Application")
               Set owbook = oexcel.Workbooks.Add
               Set osheet = owbook.Worksheets(1)
               osheet.Name = "COMPORTAMIENTO"
               Screen.MousePointer = vbHourglass
               iFila = 1
               ifila2 = 1
               icol2 = 1
               iCol = 1
               'MsgBox var_cadena
               For i = 0 To rsaux1.Fields.Count - 1
                   osheet.Cells(iFila, i + 1) = rsaux1.Fields(i).Name
                   osheet.Cells(iFila, i + 1).Font.Bold = True
               Next
               iFila = iFila + 1
                
               With osheet
                    ' carga los registros del recordset
                    .Cells(iFila, iCol).CopyFromRecordset rsaux1
                    'oExcel.Columns(1).Select
                    'oExcel.Selection.NumberFormat = "#,##0"
                    'oExcel.Columns(1).Select
                    'oExcel.Selection.Font.Color = vbRed
                    .Columns.AutoFit ' ajusta el ancho de las columnas
               End With
               owbook.SaveAs "c:\reportessid\COMPORTAMIENTO_" + Replace(Me.lv_grafica.selectedItem, " ", "_") + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               oexcel.Visible = True
               Set oexcel = Nothing
               Screen.MousePointer = vbDefault
               rsaux1.Close
               rsaux1.Open "DELETE FROM TB_TEMP_ORACLE_LECTURA_USUARIOS_HORA_HORA where inte_Tem_Consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         
            Else
               MsgBox "No existe resultados para el usuario seleccionado", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "No se a seleccionado un usuario", vbOKOnly, "ATENCION"
         End If
      '
      Else
         MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyCode = 118 Then
      If Me.lv_grafica.ListItems.Count > 0 Then
         var_usuario_h_x_h = Me.lv_grafica.selectedItem.SubItems(10)
         frmoracle_periodo_grafica_aprendizaje.Show 1
      End If
   End If
End Sub

Private Sub Timer1_Timer()
   Dim var_total As Double
   Dim list_item As ListItem
   If IsDate(Me.txt_fecha) Then
      If var_tipo_reporte = 1 Then
         var_total = 0
         var_cadena = "SELECT dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_ORACLE_LECTURA_USUARIOS.H_0_1, dbo.TB_ORACLE_LECTURA_USUARIOS.H_1_2, dbo.TB_ORACLE_LECTURA_USUARIOS.H_2_3, dbo.TB_ORACLE_LECTURA_USUARIOS.H_3_4, dbo.TB_ORACLE_LECTURA_USUARIOS.H_4_5, dbo.TB_ORACLE_LECTURA_USUARIOS.H_5_6, dbo.TB_ORACLE_LECTURA_USUARIOS.H_6_7, dbo.TB_ORACLE_LECTURA_USUARIOS.H_7_8, dbo.TB_ORACLE_LECTURA_USUARIOS.H_8_9, dbo.TB_ORACLE_LECTURA_USUARIOS.H_9_10, dbo.TB_ORACLE_LECTURA_USUARIOS.H_10_11, dbo.TB_ORACLE_LECTURA_USUARIOS.H_11_12, dbo.TB_ORACLE_LECTURA_USUARIOS.H_12_13, dbo.TB_ORACLE_LECTURA_USUARIOS.H_13_14, dbo.TB_ORACLE_LECTURA_USUARIOS.H_14_15, dbo.TB_ORACLE_LECTURA_USUARIOS.H_15_16, dbo.TB_ORACLE_LECTURA_USUARIOS.H_16_17, dbo.TB_ORACLE_LECTURA_USUARIOS.H_17_18, dbo.TB_ORACLE_LECTURA_USUARIOS.H_18_19, dbo.TB_ORACLE_LECTURA_USUARIOS.H_19_20, dbo.TB_ORACLE_LECTURA_USUARIOS.H_20_21, dbo.TB_ORACLE_LECTURA_USUARIOS.H_21_22, "
         var_cadena = var_cadena + " dbo.TB_ORACLE_LECTURA_USUARIOS.H_22_23 , dbo.TB_ORACLE_LECTURA_USUARIOS.H_23_24, dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO FROM  dbo.TB_ORACLE_LECTURA_USUARIOS INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID WHERE (dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA = '" + Me.txt_fecha + "') AND (dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO = '" + var_clave_usuario_global + "')"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         Me.lv_grafica.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_grafica.ListItems.Add(, , IIf(IsNull(rs!vcha_usu_nombre), "", rs!vcha_usu_nombre) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos))
               list_item.SubItems(1) = Format(IIf(IsNull(rs!H_23_24), 0, rs!H_23_24), "###,###,##0")
               list_item.SubItems(2) = Format(IIf(IsNull(rs!H_0_1), 0, rs!H_0_1), "###,###,##0")
               list_item.SubItems(3) = Format(IIf(IsNull(rs!H_1_2), 0, rs!H_1_2), "###,###,##0")
               list_item.SubItems(4) = Format(IIf(IsNull(rs!H_2_3), 0, rs!H_2_3), "###,###,##0")
               list_item.SubItems(5) = Format(IIf(IsNull(rs!H_3_4), 0, rs!H_3_4), "###,###,##0")
               list_item.SubItems(6) = Format(IIf(IsNull(rs!H_4_5), 0, rs!H_4_5), "###,###,##0")
               list_item.SubItems(7) = Format(IIf(IsNull(rs!H_5_6), 0, rs!H_5_6), "###,###,##0")
               list_item.SubItems(8) = Format(IIf(IsNull(rs!H_6_7), 0, rs!H_6_7), "###,###,##0")
               list_item.SubItems(9) = Format(IIf(IsNull(rs!H_23_24), 0, rs!H_23_24) + IIf(IsNull(rs!H_0_1), 0, rs!H_0_1) + IIf(IsNull(rs!H_1_2), 0, rs!H_1_2) + IIf(IsNull(rs!H_2_3), 0, rs!H_2_3) + IIf(IsNull(rs!H_3_4), 0, rs!H_3_4) + IIf(IsNull(rs!H_4_5), 0, rs!H_4_5) + IIf(IsNull(rs!H_5_6), 0, rs!H_5_6) + IIf(IsNull(rs!H_6_7), 0, rs!H_6_7), "###,###,##0")
               rs.MoveNext
         Wend
         rsaux.Open "SELECT dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_ORACLE_LECTURA_USUARIOS.H_23_24, dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO FROM  dbo.TB_ORACLE_LECTURA_USUARIOS INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID WHERE (dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA = '16/05/2012') "
         If rs.RecordCount > 0 Then
            rs.MoveFirst
         End If
         Me.lv_grafica_2.ListItems.Clear
         While Not rs.EOF
               Set list_item = lv_grafica_2.ListItems.Add(, , IIf(IsNull(rs!vcha_usu_nombre), "", rs!vcha_usu_nombre) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos))
               list_item.SubItems(1) = Format(IIf(IsNull(rs!H_7_8), 0, rs!H_7_8), "###,###,##0")
               list_item.SubItems(2) = Format(IIf(IsNull(rs!H_8_9), 0, rs!H_8_9), "###,###,##0")
               list_item.SubItems(3) = Format(IIf(IsNull(rs!H_9_10), 0, rs!H_9_10), "###,###,##0")
               list_item.SubItems(4) = Format(IIf(IsNull(rs!H_10_11), 0, rs!H_10_11), "###,###,##0")
               list_item.SubItems(5) = Format(IIf(IsNull(rs!H_11_12), 0, rs!H_11_12), "###,###,##0")
               list_item.SubItems(6) = Format(IIf(IsNull(rs!H_12_13), 0, rs!H_12_13), "###,###,##0")
               list_item.SubItems(7) = Format(IIf(IsNull(rs!H_13_14), 0, rs!H_13_14), "###,###,##0")
               list_item.SubItems(8) = Format(IIf(IsNull(rs!H_14_15), 0, rs!H_14_15), "###,###,##0")
               list_item.SubItems(9) = Format(IIf(IsNull(rs!H_15_16), 0, rs!H_15_16), "###,###,##0")
               list_item.SubItems(10) = Format(IIf(IsNull(rs!H_7_8), 0, rs!H_7_8) + IIf(IsNull(rs!H_8_9), 0, rs!H_8_9) + IIf(IsNull(rs!H_9_10), 0, rs!H_9_10) + IIf(IsNull(rs!H_10_11), 0, rs!H_10_11) + IIf(IsNull(rs!H_11_12), 0, rs!H_11_12) + IIf(IsNull(rs!H_12_13), 0, rs!H_12_13) + IIf(IsNull(rs!H_13_14), 0, rs!H_13_14) + IIf(IsNull(rs!H_14_15), 0, rs!H_14_15) + IIf(IsNull(rs!H_15_16), 0, rs!H_15_16), "###,###,##0")
               rs.MoveNext
         Wend
      
      
      
         
         rs.Close
      
      
      
      
      
      Else
         'var_cadena = "SELECT dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_ORACLE_LECTURA_USUARIOS.H_0_1, dbo.TB_ORACLE_LECTURA_USUARIOS.H_1_2, dbo.TB_ORACLE_LECTURA_USUARIOS.H_2_3, dbo.TB_ORACLE_LECTURA_USUARIOS.H_3_4, dbo.TB_ORACLE_LECTURA_USUARIOS.H_4_5, dbo.TB_ORACLE_LECTURA_USUARIOS.H_5_6, dbo.TB_ORACLE_LECTURA_USUARIOS.H_6_7, dbo.TB_ORACLE_LECTURA_USUARIOS.H_7_8, dbo.TB_ORACLE_LECTURA_USUARIOS.H_8_9, dbo.TB_ORACLE_LECTURA_USUARIOS.H_9_10, dbo.TB_ORACLE_LECTURA_USUARIOS.H_10_11, dbo.TB_ORACLE_LECTURA_USUARIOS.H_11_12, dbo.TB_ORACLE_LECTURA_USUARIOS.H_12_13, dbo.TB_ORACLE_LECTURA_USUARIOS.H_13_14, dbo.TB_ORACLE_LECTURA_USUARIOS.H_14_15, dbo.TB_ORACLE_LECTURA_USUARIOS.H_15_16, dbo.TB_ORACLE_LECTURA_USUARIOS.H_16_17, dbo.TB_ORACLE_LECTURA_USUARIOS.H_17_18, dbo.TB_ORACLE_LECTURA_USUARIOS.H_18_19, dbo.TB_ORACLE_LECTURA_USUARIOS.H_19_20, dbo.TB_ORACLE_LECTURA_USUARIOS.H_20_21, dbo.TB_ORACLE_LECTURA_USUARIOS.H_21_22, "
         'var_cadena = var_cadena + " dbo.TB_ORACLE_LECTURA_USUARIOS.H_22_23 , dbo.TB_ORACLE_LECTURA_USUARIOS.H_23_24, dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO FROM  dbo.TB_ORACLE_LECTURA_USUARIOS INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID WHERE (dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA = '" + Me.txt_fecha + "') "
         var_cadena = "SELECT dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID, "
         var_cadena = var_cadena + " dbo.TB_ORACLE_LECTURA_USUARIOS.H_0_1 , dbo.TB_ORACLE_LECTURA_USUARIOS.H_1_2, dbo.TB_ORACLE_LECTURA_USUARIOS.H_2_3, dbo.TB_ORACLE_LECTURA_USUARIOS.H_3_4, dbo.TB_ORACLE_LECTURA_USUARIOS.H_4_5, dbo.TB_ORACLE_LECTURA_USUARIOS.H_5_6, dbo.TB_ORACLE_LECTURA_USUARIOS.H_6_7, dbo.TB_ORACLE_LECTURA_USUARIOS.H_7_8, dbo.TB_ORACLE_LECTURA_USUARIOS.H_8_9, dbo.TB_ORACLE_LECTURA_USUARIOS.H_9_10, dbo.TB_ORACLE_LECTURA_USUARIOS.H_10_11, dbo.TB_ORACLE_LECTURA_USUARIOS.H_11_12, dbo.TB_ORACLE_LECTURA_USUARIOS.H_12_13, dbo.TB_ORACLE_LECTURA_USUARIOS.H_13_14, dbo.TB_ORACLE_LECTURA_USUARIOS.H_14_15, dbo.TB_ORACLE_LECTURA_USUARIOS.H_15_16, dbo.TB_ORACLE_LECTURA_USUARIOS.H_16_17, dbo.TB_ORACLE_LECTURA_USUARIOS.H_17_18, dbo.TB_ORACLE_LECTURA_USUARIOS.H_18_19, dbo.TB_ORACLE_LECTURA_USUARIOS.H_19_20, dbo.TB_ORACLE_LECTURA_USUARIOS.H_20_21, dbo.TB_ORACLE_LECTURA_USUARIOS.H_21_22, dbo.TB_ORACLE_LECTURA_USUARIOS.H_22_23, dbo.TB_ORACLE_LECTURA_USUARIOS.H_23_24, "
         var_cadena = var_cadena + " dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO , dbo.VW_ORACLE_RELACIONES_USUARIOS_UNIDADES.VCHA_UOR_UNIDAD_ID FROM  dbo.TB_ORACLE_LECTURA_USUARIOS INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID INNER JOIN dbo.VW_ORACLE_RELACIONES_USUARIOS_UNIDADES ON dbo.Tb_usuarios.VCHA_USU_USUARIO_ID = dbo.VW_ORACLE_RELACIONES_USUARIOS_UNIDADES.VCHA_USU_USUARIO_ID WHERE (dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA = '" + Me.txt_fecha + "') AND (dbo.VW_ORACLE_RELACIONES_USUARIOS_UNIDADES.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "') "
         'MsgBox var_cadena
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         Me.lv_grafica.ListItems.Clear
         Me.lv_grafica.ListItems.Clear
         var_fecha_anterior = CStr(CDate(Me.txt_fecha) - 1)
         
         While Not rs.EOF
               var_suma = IIf(IsNull(rs!H_23_24), 0, rs!H_23_24) + IIf(IsNull(rs!H_0_1), 0, rs!H_0_1) + IIf(IsNull(rs!H_1_2), 0, rs!H_1_2) + IIf(IsNull(rs!H_2_3), 0, rs!H_2_3) + IIf(IsNull(rs!H_3_4), 0, rs!H_3_4) + IIf(IsNull(rs!H_4_5), 0, rs!H_4_5) + IIf(IsNull(rs!H_5_6), 0, rs!H_5_6) + IIf(IsNull(rs!H_6_7), 0, rs!H_6_7)
              
               If var_suma > 0 Then
                  Set list_item = lv_grafica.ListItems.Add(, , IIf(IsNull(rs!vcha_usu_nombre), "", rs!vcha_usu_nombre) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos))
                  'list_item.SubItems(1) = Format(IIf(IsNull(rs!H_23_24), 0, rs!H_23_24), "###,###,##0")
                  list_item.SubItems(1) = Format(0, "###,###,##0")
                  list_item.SubItems(2) = Format(IIf(IsNull(rs!H_0_1), 0, rs!H_0_1), "###,###,##0")
                  list_item.SubItems(3) = Format(IIf(IsNull(rs!H_1_2), 0, rs!H_1_2), "###,###,##0")
                  list_item.SubItems(4) = Format(IIf(IsNull(rs!H_2_3), 0, rs!H_2_3), "###,###,##0")
                  list_item.SubItems(5) = Format(IIf(IsNull(rs!H_3_4), 0, rs!H_3_4), "###,###,##0")
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!H_4_5), 0, rs!H_4_5), "###,###,##0")
                  list_item.SubItems(7) = Format(IIf(IsNull(rs!H_5_6), 0, rs!H_5_6), "###,###,##0")
                  list_item.SubItems(8) = Format(IIf(IsNull(rs!H_6_7), 0, rs!H_6_7), "###,###,##0")
                  'list_item.SubItems(9) = Format(IIf(IsNull(rs!H_23_24), 0, rs!H_23_24) + IIf(IsNull(rs!H_0_1), 0, rs!H_0_1) + IIf(IsNull(rs!H_1_2), 0, rs!H_1_2) + IIf(IsNull(rs!H_2_3), 0, rs!H_2_3) + IIf(IsNull(rs!H_3_4), 0, rs!H_3_4) + IIf(IsNull(rs!H_4_5), 0, rs!H_4_5) + IIf(IsNull(rs!H_5_6), 0, rs!H_5_6) + IIf(IsNull(rs!H_6_7), 0, rs!H_6_7), "###,###,##0")
                  list_item.SubItems(9) = Format(0 + IIf(IsNull(rs!H_0_1), 0, rs!H_0_1) + IIf(IsNull(rs!H_1_2), 0, rs!H_1_2) + IIf(IsNull(rs!H_2_3), 0, rs!H_2_3) + IIf(IsNull(rs!H_3_4), 0, rs!H_3_4) + IIf(IsNull(rs!H_4_5), 0, rs!H_4_5) + IIf(IsNull(rs!H_5_6), 0, rs!H_5_6) + IIf(IsNull(rs!H_6_7), 0, rs!H_6_7), "###,###,##0")
                  list_item.SubItems(10) = rs!vcha_usu_usuario_id
               End If
               rs.MoveNext
         Wend
         
         'var_cadena = "SELECT dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_ORACLE_LECTURA_USUARIOS.H_23_24, dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO FROM  dbo.TB_ORACLE_LECTURA_USUARIOS INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID WHERE (dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA = '" + var_fecha_anterior + "') AND dbo.TB_ORACLE_LECTURA_USUARIOS.H_23_24 > 0"
         var_cadena = "SELECT dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA, dbo.TB_USUARIOS.VCHA_USU_NOMBRE, dbo.TB_USUARIOS.VCHA_USU_APELLIDOS, dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID,dbo.TB_ORACLE_LECTURA_USUARIOS.H_23_24, dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO, dbo.VW_ORACLE_RELACIONES_USUARIOS_UNIDADES.VCHA_UOR_UNIDAD_ID FROM  dbo.TB_ORACLE_LECTURA_USUARIOS INNER JOIN dbo.TB_USUARIOS ON dbo.TB_ORACLE_LECTURA_USUARIOS.USUARIO = dbo.TB_USUARIOS.VCHA_USU_USUARIO_ID INNER JOIN dbo.VW_ORACLE_RELACIONES_USUARIOS_UNIDADES ON dbo.Tb_usuarios.VCHA_USU_USUARIO_ID = dbo.VW_ORACLE_RELACIONES_USUARIOS_UNIDADES.VCHA_USU_USUARIO_ID WHERE (dbo.TB_ORACLE_LECTURA_USUARIOS.FECHA = '" + var_fecha_anterior + "') AND (dbo.TB_ORACLE_LECTURA_USUARIOS.H_23_24 > 0) AND (dbo.VW_ORACLE_RELACIONES_USUARIOS_UNIDADES.VCHA_UOR_UNIDAD_ID = '" + var_unidad_organizacional + "')"
         rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               var_encontro = 0
               For var_j = 1 To lv_grafica.ListItems.Count
                   lv_grafica.ListItems.Item(var_j).Selected = True
                   If Me.lv_grafica.selectedItem = IIf(IsNull(rsaux!vcha_usu_nombre), "", rsaux!vcha_usu_nombre) + " " + IIf(IsNull(rsaux!vcha_usu_apellidos), "", rsaux!vcha_usu_apellidos) Then
                      Me.lv_grafica.selectedItem.SubItems(1) = Format(IIf(IsNull(rsaux!H_23_24), 0, rsaux!H_23_24), "###,###,##0")
                      Me.lv_grafica.selectedItem.SubItems(9) = Format(CDbl(Me.lv_grafica.selectedItem.SubItems(9)) + IIf(IsNull(rsaux!H_23_24), 0, rsaux!H_23_24), "###,###,##0")
                      Me.lv_grafica.selectedItem.SubItems(10) = rsaux!vcha_usu_usuario_id
                      var_encontro = 1
                   End If
               Next var_j
               If var_encontro = 0 Then
                  Set list_item = lv_grafica.ListItems.Add(, , IIf(IsNull(rs!vcha_usu_nombre), "", rsaux!vcha_usu_nombre) + " " + IIf(IsNull(rsaux!vcha_usu_apellidos), "", rsaux!vcha_usu_apellidos))
                  list_item.SubItems(1) = Format(IIf(IsNull(rsaux!H_23_24), 0, rsaux!H_23_24), "###,###,##0")
                  list_item.SubItems(2) = Format(0, "###,###,##0")
                  list_item.SubItems(3) = Format(0, "###,###,##0")
                  list_item.SubItems(4) = Format(0, "###,###,##0")
                  list_item.SubItems(5) = Format(0, "###,###,##0")
                  list_item.SubItems(6) = Format(0, "###,###,##0")
                  list_item.SubItems(7) = Format(0, "###,###,##0")
                  list_item.SubItems(8) = Format(0, "###,###,##0")
                  list_item.SubItems(9) = Format(0, "###,###,##0")
                  'list_item.SubItems(10) = rs!VCHA_USU_USUARIO_ID
               End If
               rsaux.MoveNext
         Wend
         rsaux.Close
         var_total = 0
         If Me.lv_grafica.ListItems.Count > 0 Then
            VAR_1 = 0
            VAR_2 = 0
            VAR_3 = 0
            VAR_4 = 0
            VAR_5 = 0
            VAR_6 = 0
            VAR_7 = 0
            var_8 = 0
            var_9 = 0
            For var_j = 1 To Me.lv_grafica.ListItems.Count
                Me.lv_grafica.ListItems.Item(var_j).Selected = True
                VAR_1 = VAR_1 + CDbl(Me.lv_grafica.selectedItem.SubItems(1))
                VAR_2 = VAR_2 + CDbl(Me.lv_grafica.selectedItem.SubItems(2))
                VAR_3 = VAR_3 + CDbl(Me.lv_grafica.selectedItem.SubItems(3))
                VAR_4 = VAR_4 + CDbl(Me.lv_grafica.selectedItem.SubItems(4))
                VAR_5 = VAR_5 + CDbl(Me.lv_grafica.selectedItem.SubItems(5))
                VAR_6 = VAR_6 + CDbl(Me.lv_grafica.selectedItem.SubItems(6))
                VAR_7 = VAR_7 + CDbl(Me.lv_grafica.selectedItem.SubItems(7))
                var_8 = var_8 + CDbl(Me.lv_grafica.selectedItem.SubItems(8))
                var_9 = var_9 + CDbl(Me.lv_grafica.selectedItem.SubItems(9))
            Next var_j
            var_total = var_total + var_9
            Set list_item = lv_grafica.ListItems.Add(, , "                                     TOTAL")
            
            list_item.SubItems(1) = Format(VAR_1, "###,###,##0")
            list_item.SubItems(2) = Format(VAR_2, "###,###,##0")
            list_item.SubItems(3) = Format(VAR_3, "###,###,##0")
            list_item.SubItems(4) = Format(VAR_4, "###,###,##0")
            list_item.SubItems(5) = Format(VAR_5, "###,###,##0")
            list_item.SubItems(6) = Format(VAR_6, "###,###,##0")
            list_item.SubItems(7) = Format(VAR_7, "###,###,##0")
            list_item.SubItems(8) = Format(var_8, "###,###,##0")
            list_item.SubItems(9) = Format(var_9, "###,###,##0")
            Me.lv_grafica.ListItems.Item(var_j).Bold = True
            Me.lv_grafica.ListItems.Item(var_j).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(7).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(8).ForeColor = &HFF&
            Me.lv_grafica.ListItems.Item(var_j).ListSubItems(9).ForeColor = &HFF&
         End If
      
      
         
         If rs.RecordCount > 0 Then
            rs.MoveFirst
         End If
         Me.lv_grafica_2.ListItems.Clear
         While Not rs.EOF
               var_suma = IIf(IsNull(rs!H_7_8), 0, rs!H_7_8) + IIf(IsNull(rs!H_8_9), 0, rs!H_8_9) + IIf(IsNull(rs!H_9_10), 0, rs!H_9_10) + IIf(IsNull(rs!H_10_11), 0, rs!H_10_11) + IIf(IsNull(rs!H_11_12), 0, rs!H_11_12) + IIf(IsNull(rs!H_12_13), 0, rs!H_12_13) + IIf(IsNull(rs!H_13_14), 0, rs!H_13_14) + IIf(IsNull(rs!H_14_15), 0, rs!H_14_15)
               If var_suma > 0 Then
                  Set list_item = lv_grafica_2.ListItems.Add(, , IIf(IsNull(rs!vcha_usu_nombre), "", rs!vcha_usu_nombre) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos))
                  list_item.SubItems(1) = Format(IIf(IsNull(rs!H_7_8), 0, rs!H_7_8), "###,###,##0")
                  list_item.SubItems(2) = Format(IIf(IsNull(rs!H_8_9), 0, rs!H_8_9), "###,###,##0")
                  list_item.SubItems(3) = Format(IIf(IsNull(rs!H_9_10), 0, rs!H_9_10), "###,###,##0")
                  list_item.SubItems(4) = Format(IIf(IsNull(rs!H_10_11), 0, rs!H_10_11), "###,###,##0")
                  list_item.SubItems(5) = Format(IIf(IsNull(rs!H_11_12), 0, rs!H_11_12), "###,###,##0")
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!H_12_13), 0, rs!H_12_13), "###,###,##0")
                  list_item.SubItems(7) = Format(IIf(IsNull(rs!H_13_14), 0, rs!H_13_14), "###,###,##0")
                  list_item.SubItems(8) = Format(IIf(IsNull(rs!H_14_15), 0, rs!H_14_15), "###,###,##0")
                  list_item.SubItems(9) = Format(IIf(IsNull(rs!H_7_8), 0, rs!H_7_8) + IIf(IsNull(rs!H_8_9), 0, rs!H_8_9) + IIf(IsNull(rs!H_9_10), 0, rs!H_9_10) + IIf(IsNull(rs!H_10_11), 0, rs!H_10_11) + IIf(IsNull(rs!H_11_12), 0, rs!H_11_12) + IIf(IsNull(rs!H_12_13), 0, rs!H_12_13) + IIf(IsNull(rs!H_13_14), 0, rs!H_13_14) + IIf(IsNull(rs!H_14_15), 0, rs!H_14_15), "###,###,##0")
                  list_item.SubItems(10) = rs!vcha_usu_usuario_id
               End If
               rs.MoveNext
         Wend
         If Me.lv_grafica_2.ListItems.Count > 0 Then
            VAR_1 = 0
            VAR_2 = 0
            VAR_3 = 0
            VAR_4 = 0
            VAR_5 = 0
            VAR_6 = 0
            VAR_7 = 0
            var_8 = 0
            var_9 = 0
            For var_j = 1 To Me.lv_grafica_2.ListItems.Count
                Me.lv_grafica_2.ListItems.Item(var_j).Selected = True
                VAR_1 = VAR_1 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(1))
                VAR_2 = VAR_2 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(2))
                VAR_3 = VAR_3 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(3))
                VAR_4 = VAR_4 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(4))
                VAR_5 = VAR_5 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(5))
                VAR_6 = VAR_6 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(6))
                VAR_7 = VAR_7 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(7))
                var_8 = var_8 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(8))
                var_9 = var_9 + CDbl(Me.lv_grafica_2.selectedItem.SubItems(9))
            Next var_j
            var_total = var_total + var_9
            Set list_item = lv_grafica_2.ListItems.Add(, , "                                     TOTAL")
            
            list_item.SubItems(1) = Format(VAR_1, "###,###,##0")
            list_item.SubItems(2) = Format(VAR_2, "###,###,##0")
            list_item.SubItems(3) = Format(VAR_3, "###,###,##0")
            list_item.SubItems(4) = Format(VAR_4, "###,###,##0")
            list_item.SubItems(5) = Format(VAR_5, "###,###,##0")
            list_item.SubItems(6) = Format(VAR_6, "###,###,##0")
            list_item.SubItems(7) = Format(VAR_7, "###,###,##0")
            list_item.SubItems(8) = Format(var_8, "###,###,##0")
            list_item.SubItems(9) = Format(var_9, "###,###,##0")
            Me.lv_grafica_2.ListItems.Item(var_j).Bold = True
            Me.lv_grafica_2.ListItems.Item(var_j).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(7).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(8).ForeColor = &HFF&
            Me.lv_grafica_2.ListItems.Item(var_j).ListSubItems(9).ForeColor = &HFF&
         End If
         If rs.RecordCount > 0 Then
            rs.MoveFirst
         End If
         Me.lv_grafica_3.ListItems.Clear
         While Not rs.EOF
               var_suma = IIf(IsNull(rs!H_15_16), 0, rs!H_15_16) + IIf(IsNull(rs!H_16_17), 0, rs!H_16_17) + IIf(IsNull(rs!H_17_18), 0, rs!H_17_18) + IIf(IsNull(rs!H_18_19), 0, rs!H_18_19) + IIf(IsNull(rs!H_19_20), 0, rs!H_19_20) + IIf(IsNull(rs!H_20_21), 0, rs!H_20_21) + IIf(IsNull(rs!H_21_22), 0, rs!H_21_22) + IIf(IsNull(rs!H_22_23), 0, rs!H_22_23)
               If var_suma > 0 Then
                  Set list_item = lv_grafica_3.ListItems.Add(, , IIf(IsNull(rs!vcha_usu_nombre), "", rs!vcha_usu_nombre) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos))
                  list_item.SubItems(1) = Format(IIf(IsNull(rs!H_15_16), 0, rs!H_15_16), "###,###,##0")
                  list_item.SubItems(2) = Format(IIf(IsNull(rs!H_16_17), 0, rs!H_16_17), "###,###,##0")
                  list_item.SubItems(3) = Format(IIf(IsNull(rs!H_17_18), 0, rs!H_17_18), "###,###,##0")
                  list_item.SubItems(4) = Format(IIf(IsNull(rs!H_18_19), 0, rs!H_18_19), "###,###,##0")
                  list_item.SubItems(5) = Format(IIf(IsNull(rs!H_19_20), 0, rs!H_19_20), "###,###,##0")
                  list_item.SubItems(6) = Format(IIf(IsNull(rs!H_20_21), 0, rs!H_20_21), "###,###,##0")
                  list_item.SubItems(7) = Format(IIf(IsNull(rs!H_21_22), 0, rs!H_21_22), "###,###,##0")
                  list_item.SubItems(8) = Format(IIf(IsNull(rs!H_22_23), 0, rs!H_22_23), "###,###,##0")
                  list_item.SubItems(9) = Format(IIf(IsNull(rs!H_15_16), 0, rs!H_15_16) + IIf(IsNull(rs!H_16_17), 0, rs!H_16_17) + IIf(IsNull(rs!H_17_18), 0, rs!H_17_18) + IIf(IsNull(rs!H_18_19), 0, rs!H_18_19) + IIf(IsNull(rs!H_19_20), 0, rs!H_19_20) + IIf(IsNull(rs!H_20_21), 0, rs!H_20_21) + IIf(IsNull(rs!H_21_22), 0, rs!H_21_22) + IIf(IsNull(rs!H_22_23), 0, rs!H_22_23), "###,###,##0")
                  list_item.SubItems(10) = rs!vcha_usu_usuario_id
               End If
               rs.MoveNext
         Wend
         If Me.lv_grafica_3.ListItems.Count > 0 Then
            VAR_1 = 0
            VAR_2 = 0
            VAR_3 = 0
            VAR_4 = 0
            VAR_5 = 0
            VAR_6 = 0
            VAR_7 = 0
            var_8 = 0
            var_9 = 0
            For var_j = 1 To Me.lv_grafica_3.ListItems.Count
                Me.lv_grafica_3.ListItems.Item(var_j).Selected = True
                VAR_1 = VAR_1 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(1))
                VAR_2 = VAR_2 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(2))
                VAR_3 = VAR_3 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(3))
                VAR_4 = VAR_4 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(4))
                VAR_5 = VAR_5 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(5))
                VAR_6 = VAR_6 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(6))
                VAR_7 = VAR_7 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(7))
                var_8 = var_8 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(8))
                var_9 = var_9 + CDbl(Me.lv_grafica_3.selectedItem.SubItems(9))
                
            Next var_j
            var_total = var_total + var_9
            
            Set list_item = lv_grafica_3.ListItems.Add(, , "                                     TOTAL")
         
            list_item.SubItems(1) = Format(VAR_1, "###,###,##0")
            list_item.SubItems(2) = Format(VAR_2, "###,###,##0")
            list_item.SubItems(3) = Format(VAR_3, "###,###,##0")
            list_item.SubItems(4) = Format(VAR_4, "###,###,##0")
            list_item.SubItems(5) = Format(VAR_5, "###,###,##0")
            list_item.SubItems(6) = Format(VAR_6, "###,###,##0")
            list_item.SubItems(7) = Format(VAR_7, "###,###,##0")
            list_item.SubItems(8) = Format(var_8, "###,###,##0")
            list_item.SubItems(9) = Format(var_9, "###,###,##0")
            Me.lv_grafica_3.ListItems.Item(var_j).Bold = True
            Me.lv_grafica_3.ListItems.Item(var_j).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(7).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(8).ForeColor = &HFF&
            Me.lv_grafica_3.ListItems.Item(var_j).ListSubItems(9).ForeColor = &HFF&
         End If
         Me.lbl_total = Format(var_total, "###,###,##0")
   
         rs.Close
      End If
   Else
      MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fecha = var_fecha_general
   End If
End Sub
