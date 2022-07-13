VERSION 5.00
Begin VB.Form frmreporte_relacion_cobranza_auditoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Folio relación cobranza"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   105
      TabIndex        =   5
      Top             =   1260
      Width           =   4245
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   375
         Width           =   420
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3975
      Picture         =   "frmreporte_relacion_cobranza_auditoria.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmreporte_relacion_cobranza_auditoria.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   75
      TabIndex        =   4
      Top             =   360
      Width           =   4305
   End
   Begin VB.Frame Frame1 
      Caption         =   " Folio "
      Height          =   735
      Left            =   105
      TabIndex        =   3
      Top             =   495
      Width           =   4230
      Begin VB.TextBox txt_folio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1020
         TabIndex        =   0
         Top             =   195
         Width           =   2280
      End
   End
End
Attribute VB_Name = "frmreporte_relacion_cobranza_auditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
  ' On Error GoTo salir:
   If Trim(Me.txt_folio) <> "" Then
      If rsaux.State = 1 Then
         rsaux.Close
      End If
      rsaux.Open "SELECT * FROM TB_RELACION_COBRANZA with (nolock) WHERE VCHA_RCO_FOLIO = '" + Me.txt_folio + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Set reporte = appl.OpenReport(App.Path + "\REP_RELACION_COBRANZA_AUDITORIA.rpt")
         reporte.RecordSelectionFormula = "{VW_RELACION_COBRANZA_AUDITORIA.VCHA_RCO_FOLIO} = '" + Me.txt_folio + "'"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\relacion_cobranza_" + Trim(Me.txt_folio) + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
      Else
         MsgBox "La relación de cobranza seleccionada no existe", vbOKOnly, "ATENCION"
      End If
      rsaux.Close
   Else
      If IsDate(Me.txt_fin) Then
         If IsDate(Me.txt_inicio) Then
            dia = CStr(Day(CDate(Me.txt_inicio)))
            mes = CStr(Month(CDate(Me.txt_inicio)))
            año = CStr(Year(CDate(Me.txt_inicio)))
            If Len(dia) = 1 Then
               dia = "0" + dia
            End If
            If Len(mes) = 1 Then
               mes = "0" + mes
            End If
            If Len(año) = 1 Then
               año = "200" + año
            Else
               If Len(año) = 2 Then
                  año = "20" + año
               End If
            End If
            var_fecha_inicio = "{d '" + año + "-" + mes + "-" + dia + "'}"
            VAR_FECHA_INICIO_CRYSTAL = año + "," + mes + "," + dia + ",00,00,00"
         
            dia = CStr(Day(CDate(Me.txt_fin)))
            mes = CStr(Month(CDate(Me.txt_fin)))
            año = CStr(Year(CDate(Me.txt_fin)))
            If Len(dia) = 1 Then
               dia = "0" + dia
            End If
            If Len(mes) = 1 Then
               mes = "0" + mes
            End If
            If Len(año) = 1 Then
               año = "200" + año
            Else
               If Len(año) = 2 Then
                  año = "20" + año
               End If
            End If
            var_fecha_fin = "{d '" + año + "-" + mes + "-" + dia + "'}"
            VAR_FECHA_FIN_CRYSTAL = año + "," + mes + "," + dia + ",23,59,59"
      
            rs.Open "select * from tb_relacion_cobranza where DTIM_RCO_FECHA_RELACION >= " + var_fecha_inicio + " and DTIM_RCO_FECHA_RELACION <= " + var_fecha_fin + " + 1 - .0000001", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
         
               Set reporte = appl.OpenReport(App.Path + "\REP_RELACION_COBRANZA_AUDITORIA.rpt")
               reporte.RecordSelectionFormula = "{VW_RELACION_COBRANZA_AUDITORIA.VCHA_RCO_FOLIO} = '" + Me.txt_folio + "'"
               reporte.RecordSelectionFormula = "{VW_RELACION_COBRANZA_AUDITORIA.DTIM_RCO_FECHA_RELACION}>= DATETIME(" + VAR_FECHA_INICIO_CRYSTAL + ") AND {VW_RELACION_COBRANZA_AUDITORIA.DTIM_RCO_FECHA_RELACION}<= DATETIME(" + VAR_FECHA_FIN_CRYSTAL + ")"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\relacion_cobranza_" + Trim(Me.txt_folio) + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + ".xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               MsgBox "Se a terminado de guardar el archivo " + archivo
            Else
               MsgBox "No existe relaciones de cobranza para la fecha seleccionada", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "Fecha inicio incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   
   Exit Sub
salir:
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   MsgBox "A surgido un error al generar el archivo, puede que este este abierto.", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 2500
   Left = 3500
   Me.txt_fin = Date
   Me.txt_inicio = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
End Sub

Private Sub txt_fin_Change()
   Me.txt_folio = ""
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
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

Private Sub txt_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_inicio_Change()
   Me.txt_folio = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
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
