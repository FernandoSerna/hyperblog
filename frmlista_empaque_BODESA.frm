VERSION 5.00
Begin VB.Form frmlista_empaque_bodesa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista empaque BODESA"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Embarque "
      Height          =   1200
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   2805
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   255
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   375
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmlista_empaque_bodesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Private Sub Form_Load()
   Me.txt_embarque = ""
   Top = 3000
   Left = 3900
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         rs.Open "SELECT * FROM TB_dETALLE_cAJAS with (nolock) WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Set reporte = appl.OpenReport(App.Path + "\rep_lista_embarque_bodesa.rpt")
            reporte.RecordSelectionFormula = "{VW_FACTURACION_TALLAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_FACTURACION_TALLAS.INTE_EMB_EMBARQUE} = " + Me.txt_embarque
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Lista de empaque de BODESA"
            frmvistasprevias.Show 1
            Set reporte = Nothing
            
            var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
            If var_si = 6 Then
               Set reporte = appl.OpenReport(App.Path + "\rep_lista_embarque_bodesa_excel.rpt")
               reporte.RecordSelectionFormula = "{VW_FACTURACION_TALLAS.VCHA_EMP_EMPRESA_ID} = '" + var_empresa + "' and {VW_FACTURACION_TALLAS.INTE_EMB_EMBARQUE} = " + Me.txt_embarque
               For ntablas = 1 To reporte.Database.Tables.Count
                  reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\reporte_lista_empaque_BODESA_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            End If
            
            

         Else
            MsgBox "Número de embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "Embarque incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

