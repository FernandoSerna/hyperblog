VERSION 5.00
Begin VB.Form frmbusqueda_documento_cargo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de documento"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3285
      Picture         =   "frmbusqueda_documento_cargo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmbusqueda_documento_cargo.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar"
      Top             =   45
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   30
      TabIndex        =   5
      Top             =   345
      Width           =   3600
   End
   Begin VB.Frame Frame1 
      Caption         =   " Documento "
      Height          =   1215
      Left            =   75
      TabIndex        =   0
      Top             =   450
      Width           =   3540
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   1770
         TabIndex        =   2
         Top             =   630
         Width           =   1605
      End
      Begin VB.ComboBox cmb_documento 
         Height          =   315
         ItemData        =   "frmbusqueda_documento_cargo.frx":073C
         Left            =   150
         List            =   "frmbusqueda_documento_cargo.frx":074F
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   300
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   713
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmbusqueda_documento_cargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_nuevo_Click()

End Sub

Private Sub cmb_documento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.Text1.SetFocus
   End If
End Sub

Private Sub cmd_buscar_Click()
 Dim var_documento As String
 Dim var_numero As Double
 If Me.cmb_documento.Text = "FACTURA" Then
    var_documento = "FA"
 Else
    If Me.cmb_documento.Text = "NOTA DE CARGO" Then
       var_documento = "NG"
    Else
       If Me.cmb_documento.Text = "CHEQUE DEVUELTO" Then
          var_documento = "CH"
       Else
          If Me.cmb_documento.Text = "CARGO" Then
             var_documento = "CR"
          Else
             If Me.cmb_documento.Text = "SUSTITUCION" Then
                var_documento = "SU"
             Else
                var_documento = ""
             End If
          End If
       End If
    End If
 End If
 If Trim(var_documento) <> "" Then
    If Trim(Text1) = "" Then
       Text1 = "0"
    End If
    var_numero = CDbl(Me.Text1)
    'MsgBox cnn.ConnectionString
    rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_CAR_TIPO_DOCUMENTO = '" + var_documento + "' AND INTE_CAR_NUMERO = " + CStr(var_numero) + " and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       Set reporte = appl.OpenReport(App.Path + "\REP_BUSQUEDA_CARGO.rpt")
       
       reporte.RecordSelectionFormula = "{VW_CARTERA_CARGOS_BUSQUEDA.VCHA_CAR_TIPO_DOCUMENTO} = '" + var_documento + "' AND {VW_CARTERA_CARGOS_BUSQUEDA.INTE_ECU_NUMERO_CARGO} = " + CStr(var_numero) + " and {VW_CARTERA_CARGOS_BUSQUEDA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
       frmvistasprevias.cr.ReportSource = reporte
       For ntablas = 1 To reporte.Database.Tables.Count
           reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
       Next ntablas
       frmvistasprevias.cr.ViewReport
       frmvistasprevias.Caption = "Abonos aplicados a un cargo"
       frmvistasprevias.Show 1
       Set reporte = Nothing
       
       var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
       If var_si = 6 Then
          Set reporte = appl.OpenReport(App.Path + "\REP_BUSQUEDA_CARGO_2.rpt")
          reporte.RecordSelectionFormula = "{VW_CARTERA_CARGOS_BUSQUEDA.VCHA_CAR_TIPO_DOCUMENTO} = '" + var_documento + "' AND {VW_CARTERA_CARGOS_BUSQUEDA.INTE_ECU_NUMERO_CARGO} = " + CStr(var_numero) + " and {VW_CARTERA_CARGOS_BUSQUEDA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "'"
          For ntablas = 1 To reporte.Database.Tables.Count
              reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
          Next ntablas
          reporte.ExportOptions.Reset
          'reporte.ExportOptions.FormatType = crEFTPortableDocFormat
          'reporte.ExportOptions.FormatType = crEFTWordForWindows
          reporte.ExportOptions.FormatType = crEFTExcel80
          reporte.ExportOptions.DestinationType = crEDTDiskFile
          reporte.ExportOptions.UseReportDateFormat = True
          reporte.ExportOptions.UseReportNumberFormat = True
          'archivo = "c:\reportessid\Reporte_abonos_a_cargo_" & Replace(Str(Date), "/", "") & "_" & Replace(Str(Time), ":", ".") & ".xls"
          archivo = "c:\reportessid\Reporte_abonos_a_cargo_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
          reporte.ExportOptions.DiskFileName = archivo
          reporte.Export False
          Set reporte = Nothing
          MsgBox "Se a terminado de guardar el archivo " + archivo
       End If
       
       
       
    Else
       MsgBox "El documento no existe", vbOKOnly, "ATENCION"
    End If
    rs.Close
 Else
    MsgBox "Documento Incorrecto", vbOKOnly, "ATENCION"
 End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim dl As Long                                 ' Valor devuelto por la función API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripción del DSN
Dim sDsnName As String                  ' Nombre del DSN

   cnn.Close
   cnn.Open var_conexion_string_distribucion

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   sDsnName = "DSN=sqlsistema"
   sDriver = "SQL Server"
   dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

   'se crea
   sDsnName = "sqlsistema"
   sDescription = "sqlsistema"
   sDriver = "SQL Server"
   sAttributes = "DSN=" & sDsnName & Chr(0)
   sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   
   Top = 2500
   Left = 4000
   Me.cmb_documento.Text = "FACTURA"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_comisiones)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_buscar.SetFocus
   End If
End Sub
