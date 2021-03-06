VERSION 5.00
Begin VB.Form frmreporte_ajustes_reempaque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Ajustes por Reempaque"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4470
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   4335
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   5
         Top             =   255
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   4
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   6
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmreporte_ajustes_reempaque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4035
      Picture         =   "frmreporte_ajustes_reempaque.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_ajustes_reempaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_clave_movimiento As String
   Dim var_fecha_inicio As String, var_fecha_fin As String
   Dim var_consecutivo As Double
   
   var_clave_movimiento = ""
   rs.Open "select isnull(vcha_mov_movimiento_id,'') from tb_movimientos where inte_mov_reempaque = 2", cnn, adOpenDynamic, adLockOptimistic
   If rs.EOF Then
      var_clave_movimiento = ""
   Else
      var_clave_movimiento = rs(0).Value
   End If
   rs.Close
   If Trim(var_clave_movimiento) <> "" Then
      If IsDate(txt_inicio) Then
         If IsDate(txt_fin) Then
            If CDate(txt_inicio) <= CDate(txt_fin) Then
               cnn.BeginTrans
               rs.Open "select max(inte_ree_consecutivo) from TB_TEMP_REEMPAQUE_AJUSTES", cnn, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
               Else
                  var_consecutivo = 0
               End If
               var_consecutivo = var_consecutivo + 1
               rs.Close
               rs.Open "Insert into TB_TEMP_REEMPAQUE_AJUSTES (INTE_REE_CONSECUTIVO, vcha_aud_usuario, vcha_aud_maquina) values (" + CStr(var_consecutivo) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               var_fecha_fin_1 = CDate(txt_fin) + 1
               
               var_dia = CStr(Day(CDate(txt_inicio)))
               var_mes = CStr(Month(CDate(txt_inicio)))
               var_a?o = CStr(Year(CDate(txt_inicio)))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_inicio = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
             
             
               var_dia = CStr(Day(var_fecha_fin_1))
               var_mes = CStr(Month(var_fecha_fin_1))
               var_a?o = CStr(Year(var_fecha_fin_1))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_fin = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
               
               rs.Open "SELECT * FROM TB_ENCABEZADO_MOVIMIENTOS WHERE VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "' AND dtim_emo_fecha >= " + var_fecha_inicio + " and dtim_emo_fecha <= " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
               
               var_fecha_fin_1 = CDate(txt_fin)
               var_dia = CStr(Day(var_fecha_fin_1))
               var_mes = CStr(Month(var_fecha_fin_1))
               var_a?o = CStr(Year(var_fecha_fin_1))
               If Len(Trim(var_dia)) = 1 Then
                  var_dia = "0" + var_dia
               End If
               If Len(Trim(var_mes)) = 1 Then
                  var_mes = "0" + var_mes
               End If
               var_fecha_fin = "{d '" + var_a?o + "-" + var_mes + "-" + var_dia + "'}"
               While Not rs.EOF
                     rsaux.Open "insert into TB_TEMP_REEMPAQUE_AJUSTES (INTE_REE_CONSECUTIVO, DTIM_REE_FECHA_INICIO, DTIM_REE_FECHA_FIN, VCHA_EMO_EMPRESA_ID, VCHA_UOR_UNIDAD_ID, VCHA_ALM_ALMACEN_ID, VCHA_MOV_MOVIMIENTO_ID, INTE_EMO_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA) VALUES (" + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + rs!VCHA_EMP_EMPRESA_ID + "', '" + rs!vcha_uor_unidad_id + "', '" + rs!VCHA_ALM_ALMACEN_ID + "', '" + rs!VCHA_MOV_MOVIMIENTO_ID + "', " + CStr(rs!INTE_EMO_NUMERO) + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "')", cnn, adOpenDynamic, adLockOptimistic
                     rs.MoveNext
               Wend
               rs.Close
               
               Set reporte = appl.OpenReport(App.Path + "\rep_ajustes_reempaque.rpt")
               reporte.RecordSelectionFormula = "{VW_REEMPAQUE_AJUSTES_PERIODOS.INTE_REE_CONSECUTIVO}=" + CStr(var_consecutivo)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de Ajustes de Reempaque"
               frmvistasprevias.Show 1
            Else
               MsgBox "La fecha de inicio debe de ser menor a la fecha final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No existe un movimiento de reempaque", vbOKOnly, "ATENCION"
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
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_reporte_ajustes_reempaque)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub
