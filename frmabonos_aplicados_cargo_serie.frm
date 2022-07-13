VERSION 5.00
Begin VB.Form frmabonos_aplicados_cargo_serie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda de documentos"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Documento "
      Height          =   1935
      Left            =   75
      TabIndex        =   7
      Top             =   420
      Width           =   3540
      Begin VB.TextBox txt_Serie 
         Height          =   360
         Left            =   1755
         TabIndex        =   1
         Top             =   645
         Width           =   1605
      End
      Begin VB.TextBox txt_numero_fin 
         Height          =   360
         Left            =   1755
         TabIndex        =   3
         Top             =   1455
         Width           =   1605
      End
      Begin VB.ComboBox cmb_documento 
         Height          =   315
         ItemData        =   "frmabonos_aplicados_cargo_serie.frx":0000
         Left            =   150
         List            =   "frmabonos_aplicados_cargo_serie.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   285
         Width           =   3240
      End
      Begin VB.TextBox txt_numero_inicio 
         Height          =   360
         Left            =   1755
         TabIndex        =   2
         Top             =   1050
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   735
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número fin:"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1545
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número inicio:"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   1140
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmabonos_aplicados_cargo_serie.frx":0044
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Buscar"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3285
      Picture         =   "frmabonos_aplicados_cargo_serie.frx":0146
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   30
      TabIndex        =   6
      Top             =   315
      Width           =   3600
   End
End
Attribute VB_Name = "frmabonos_aplicados_cargo_serie"
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
      Call pro_enfoque(KeyAscii)
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
             var_documento = ""
          End If
       End If
    End If
 End If
 If IsNumeric(Me.txt_numero_inicio) Then
    If IsNumeric(Me.txt_numero_fin) Then
       rs.Open "SELECT * FROM TB_ENCABEZADO_CARTERA WHERE VCHA_CAR_TIPO_DOCUMENTO = '" + var_documento + "' AND INTE_CAR_NUMERO >= " + CStr(Me.txt_numero_inicio) + " AND INTE_cAR_NUMERO <= " + CStr(Me.txt_numero_fin) + " and vcha_emp_empresa_id = '" + var_empresa + "' AND VCHA_sER_SERIE_ID = '" + Me.txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          Set reporte = appl.OpenReport(App.Path + "\REP_BUSQUEDA_CARGO_EXPORTA.rpt")
          reporte.RecordSelectionFormula = "{VW_CARTERA_CARGOS_BUSQUEDA.VCHA_CAR_TIPO_DOCUMENTO} = '" + var_documento + "' AND {VW_CARTERA_CARGOS_BUSQUEDA.INTE_ECU_NUMERO_CARGO} >= " + CStr(Me.txt_numero_inicio) + " AND  {VW_CARTERA_CARGOS_BUSQUEDA.INTE_ECU_NUMERO_CARGO}<= " + Me.txt_numero_fin + " and {VW_CARTERA_CARGOS_BUSQUEDA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CARGOS_BUSQUEDA.VCHA_ECU_sERIE_CARGO} = '" + Me.txt_serie + "'"
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
             Set reporte = appl.OpenReport(App.Path + "\REP_BUSQUEDA_CARGO_EXPORTA.rpt")
             reporte.RecordSelectionFormula = "{VW_CARTERA_CARGOS_BUSQUEDA.VCHA_CAR_TIPO_DOCUMENTO} = '" + var_documento + "' AND {VW_CARTERA_CARGOS_BUSQUEDA.INTE_ECU_NUMERO_CARGO} >= " + CStr(Me.txt_numero_inicio) + " AND  {VW_CARTERA_CARGOS_BUSQUEDA.INTE_ECU_NUMERO_CARGO}<= " + Me.txt_numero_fin + " and {VW_CARTERA_CARGOS_BUSQUEDA.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' AND {VW_CARTERA_CARGOS_BUSQUEDA.VCHA_ECU_sERIE_CARGO} = '" + Me.txt_serie + "'"
             For ntablas = 1 To reporte.Database.Tables.Count
                 reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
             Next ntablas
             'reporte.ExportOptions.FormatType = crEFTPortableDocFormat
             reporte.ExportOptions.FormatType = crEFTExcel80
             reporte.ExportOptions.DestinationType = crEDTDiskFile
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
       MsgBox "Número final incorrecto", vbOKOnly, "ATENCION"
    End If
 Else
    MsgBox "Número de inicio incorrecto", vbOKOnly, "ATENCION"
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
   Call activa_forma(var_activa_forma_packing_list)
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

Private Sub Text3_Change()

End Sub

Private Sub txt_numero_fin_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_numero_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub
