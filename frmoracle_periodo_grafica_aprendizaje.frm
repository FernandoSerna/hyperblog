VERSION 5.00
Begin VB.Form frmoracle_periodo_grafica_aprendizaje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Periodo"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3480
      Picture         =   "frmoracle_periodo_grafica_aprendizaje.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Picture         =   "frmoracle_periodo_grafica_aprendizaje.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   3915
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   795
         TabIndex        =   0
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   2580
         TabIndex        =   1
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2295
         TabIndex        =   5
         Top             =   375
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmoracle_periodo_grafica_aprendizaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   If IsDate(Me.txt_fecha_inicio) Then
      If IsDate(Me.txt_fecha_fin) Then
         If CDate(Me.txt_fecha_inicio) <= CDate(Me.txt_fecha_fin) Then
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_CURVA_CONOCIMIENTO", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "INSERT INTO TB_TEMP_ORACLE_CURVA_CONOCIMIENTO (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            
            var_año_str = CStr(Year(CDate(Me.txt_fecha_inicio)))
            var_dia_str = CStr(Day(CDate(Me.txt_fecha_inicio)))
            var_mes_str = CStr(Month(CDate(Me.txt_fecha_inicio)))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha_inicio = "{d '" + CStr(var_año_str) + "-" + CStr(var_mes_str) + "-" + CStr(var_dia_str) + "'}"
            
            
            var_año_str = CStr(Year(CDate(Me.txt_fecha_fin)))
            var_dia_str = CStr(Day(CDate(Me.txt_fecha_fin)))
            var_mes_str = CStr(Month(CDate(Me.txt_fecha_fin)))
            If Len(var_dia_str) = 1 Then
               var_dia_str = "0" + var_dia_str
            End If
            If Len(var_mes_str) = 1 Then
               var_mes_str = "0" + var_mes_str
            End If
            If Len(var_año_str) = 2 Then
               var_año_str = "20" + var_año_str
            End If
            var_fecha_fin = "{d '" + CStr(var_año_str) + "-" + CStr(var_mes_str) + "-" + CStr(var_dia_str) + "'}"
            
            rs.Open "execute SP_ORACLE_CURVA_APRENDIZAJE " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + var_usuario_h_x_h + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "DELETE FROM TB_TEMP_ORACLE_CURVA_CONOCIMIENTO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and usuario is null", cnn, adOpenDynamic, adLockOptimistic
            
             Set reporte = appl.OpenReport(App.Path + "\rep_curva_aprendizaje.rpt")
             var_cadena = "{VW_ORACLE_CURVA_APRENDIZAJE.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
             reporte.RecordSelectionFormula = var_cadena
             frmvistasprevias.cr.ReportSource = reporte
             For ntablas = 1 To reporte.Database.Tables.Count
                 reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
             Next ntablas
             frmvistasprevias.cr.ViewReport
             frmvistasprevias.Caption = "Curva aprendizaje"
             frmvistasprevias.Show 1
             Set reporte = Nothing
            
            
            
            
            rs.Open "DELETE FROM TB_TEMP_ORACLE_CURVA_CONOCIMIENTO WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            
            
            
         Else
            MsgBox "La fecha final debe de ser mayor o igaul a la fecha de inicio", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
         Me.txt_fecha_fin = Date
         
      End If
   Else
      MsgBox "Fecha de inicio incorrecto", vbOKOnly, "ATENCION"
      Me.txt_fecha_inicio = Date
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.txt_fecha_fin = Date
   Me.txt_fecha_inicio = Date
   
End Sub

Private Sub txt_fecha_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_fin) Then
         frmcalendario.mes = CDate(Me.txt_fecha_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_inicio) Then
         frmcalendario.mes = CDate(Me.txt_fecha_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fecha_inicio = var_fecha_general
   End If
End Sub
