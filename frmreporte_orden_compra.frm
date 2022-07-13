VERSION 5.00
Begin VB.Form frmreporte_orden_compra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de Compra"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   0
      TabIndex        =   5
      Top             =   345
      Width           =   2715
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmreporte_orden_compra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_orden_compra.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2340
      Picture         =   "frmreporte_orden_compra.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   " Orden de Compra "
      Height          =   750
      Left            =   45
      TabIndex        =   0
      Top             =   480
      Width           =   2640
      Begin VB.TextBox txt_embarque 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   2370
      End
   End
End
Attribute VB_Name = "frmreporte_orden_compra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If IsNumeric(txt_embarque) Then
      rs.Open "select * from vw_reporte_orden_compra where inte_com_numero = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Set reporte = appl.OpenReport(App.Path + "\rep_orden_compra.rpt")
         reporte.RecordSelectionFormula = "{VW_REPORTE_ORDEN_COMPRA.INTE_COM_NUMERO} = " + txt_embarque
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Orden de Compra"
         frmvistasprevias.Show 1
         Set reporte = Nothing
      
         var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            Set reporte = appl.OpenReport(App.Path + "\rep_orden_compra.rpt")
            reporte.RecordSelectionFormula = "{VW_REPORTE_ORDEN_COMPRA.INTE_COM_NUMERO} = " + txt_embarque
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            reporte.ExportOptions.FormatType = crEFTExcel80
            reporte.ExportOptions.DestinationType = crEDTDiskFile
            archivo = "c:\reportessid\Orden_Compra_" + Trim(txt_embarque) + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            reporte.ExportOptions.DiskFileName = archivo
            reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
  End If
      
      
      Else
         MsgBox "La orden de compra", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      MsgBox "Número de orden de compra incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   txt_embarque = ""
   txt_embarque.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3850
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      cmd_imprimir.SetFocus
   End If
End Sub

