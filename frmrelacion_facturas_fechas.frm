VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrelacion_facturas_fechas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relación de facturas por fecha"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2205
      Left            =   105
      TabIndex        =   11
      Top             =   -45
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1665
         Left            =   45
         TabIndex        =   12
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   2937
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Agente "
      Height          =   795
      Left            =   135
      TabIndex        =   10
      Top             =   465
      Width           =   5595
      Begin VB.TextBox txt_nombre_agente 
         Height          =   375
         Left            =   1020
         TabIndex        =   1
         Top             =   285
         Width           =   4485
      End
      Begin VB.TextBox txt_agente 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   285
         Width           =   885
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      Picture         =   "frmrelacion_facturas_fechas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmrelacion_facturas_fechas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   90
      TabIndex        =   9
      Top             =   345
      Width           =   5640
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   5625
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3525
         TabIndex        =   3
         Top             =   255
         Width           =   1095
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1515
         TabIndex        =   2
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3210
         TabIndex        =   8
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1065
         TabIndex        =   7
         Top             =   315
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmrelacion_facturas_fechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Double
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   
   'On Error GoTo salir:
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            
             var_fecha_fin_1 = CDate(txt_fin) + 1
             var_dia = CStr(Day(CDate(txt_inicio)))
             var_mes = CStr(Month(CDate(txt_inicio)))
             var_año = CStr(Year(CDate(txt_inicio)))
             
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
             
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_año = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            cnn.BeginTrans
            rs.Open "select max(isnull(INTE_TEM_CONSECUTIVO,0)) as numero from TB_TEM_REPORTE_RELACION_FACTURAS_FECHAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            var_consecutivo = var_consecutivo + 1
            rs.Close
            rsaux.Open "INSERT INTO TB_TEM_REPORTE_RELACION_FACTURAS_FECHAS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN) Values ( " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
             'rs.Open "INSERT INTO TB_TEMP_VALUACION_DEVOLUCIONES (INTE_TVD_CONSECUTIVO, DTIM_TVD_FECHA_INICIO, DTIM_TVD_FECHA_FIN, VCHA_EMP_EMPRESA_ID, VCHA_CAR_TIPO_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO, VCHA_AUD_USUARIO, VCHA_AUD_MAQUINA) select * from tb_encabezado_cartera where dtim_Car_fecha >= " + var_fecha_inicio + " and vcha_car_tipo_documento = 'NC' and VCHA_CAR_DOCUMENTO = 'DV' and dtim_car_fecha <= " + var_fecha_fin, cnn, adOpenDynamic, adLockOptimistic
              rsaux.Open "INSERT INTO TB_TEM_REPORTE_RELACION_FACTURAS_FECHAS (INTE_TEM_CONSECUTIVO, DTIM_TEM_FECHA_INICIO, DTIM_TEM_FECHA_FIN, VCHA_AGE_AGENTE_ID, VCHA_EMP_EMPRESA_ID, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO) select " + CStr(var_consecutivo) + ", " + var_fecha_inicio + ", " + var_fecha_fin + "-.00001,'" + Me.txt_agente + "', vcha_emp_empresa_id, VCHA_CAR_DOCUMENTO, VCHA_SER_SERIE_ID, INTE_CAR_NUMERO from TB_ENCABEZADO_cartera where VCHA_AGE_AGENTE_ID = '" + Me.txt_agente + "' AND VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND (CHAR_CAR_eSTATUS <> 'C' OR CHAR_CAR_ESTATUS IS NULL) AND dtim_CAR_fecha >= " + var_fecha_inicio + " and dtim_CAR_fecha <= " + var_fecha_fin + "-.00001 AND VCHA_cAR_DOCUMENTO = 'FA'", cnn, adOpenDynamic, adLockOptimistic

             var_fecha_fin_1 = CDate(txt_fin)
             var_dia = CStr(Day(var_fecha_fin_1))
             var_mes = CStr(Month(var_fecha_fin_1))
             var_año = CStr(Year(var_fecha_fin_1))
             If Len(Trim(var_dia)) = 1 Then
                var_dia = "0" + var_dia
             End If
             If Len(Trim(var_mes)) = 1 Then
                var_mes = "0" + var_mes
             End If
             var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
             
         Set reporte = appl.OpenReport(App.Path + "\rep_relacion_facturas_fechas.rpt")
         reporte.RecordSelectionFormula = "{VW_RELACION_FACTURAS_FECHAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
         frmvistasprevias.cr.ReportSource = reporte
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         frmvistasprevias.cr.ViewReport
         frmvistasprevias.Caption = "Reporte de relación de facturas"
         frmvistasprevias.Show 1
         Set reporte = Nothing
         
         rs.Open "delete from TB_TEM_REPORTE_RELACION_FACTURAS_FECHAS where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La fecha de inicio debe de ser mayor", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha Final Incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de Inicio Incorrecta", vbOKOnly, "ATENCION"
   End If
   Exit Sub
salir:
   MsgBox "A surgido un error al generar el reporte", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
End Sub


Private Sub cmd_salir_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3200
   txt_inicio = Date
   txt_fin = Date
   frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_agente = Me.lv_lista.selectedItem
      Me.txt_nombre_agente = Me.lv_lista.selectedItem.SubItems(1)
      Me.txt_agente.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.txt_agente.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "SELECT * FROM TB_aGENTES WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
      
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "Agentes"
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 4270.71
      Else
         lv_lista.ColumnHeaders(2).Width = 4499.71
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(Me.txt_agente) <> "" Then
      rs.Open "select * from tb_agentes where vcha_age_agente_id = '" + Me.txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente incorrecto", vbOKOnly, "ATENCION"
      End If
      rs.Close
   Else
      Me.txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub


Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_inicio.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub
