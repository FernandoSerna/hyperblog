VERSION 5.00
Begin VB.Form frmreporte_nivel_surtido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Nivel de Surtido"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4440
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   90
      TabIndex        =   4
      Top             =   435
      Width           =   4245
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   315
         Width           =   1140
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   1140
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3975
      Picture         =   "frmreporte_nivel_surtido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmreporte_nivel_surtido.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmreporte_nivel_surtido.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Width           =   4395
   End
End
Attribute VB_Name = "frmreporte_nivel_surtido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_tipo_mes As Integer

Private Sub cmd_imprimir_Click()
   Set TB_NIVEL_SURTIDO = New TB_NIVEL_SURTIDO
   var_añadir = TB_NIVEL_SURTIDO.Anadir(txt_fecha_inicio, txt_fecha_fin, var_clave_usuario_global, fun_NombrePc)
            Set reporte = appl.OpenReport(App.Path + "\rep_salida_vistas.rpt")
            reporte.RecordSelectionFormula = "{VW_orden_surtido_mov.inte_emo_numero} = " + Str(var_numero_folio) + " and {VW_ORDEN_SURTIDO_MOV.FLOA_SAL_CANTIDAD} > 0 and {VW_ORDEN_SURTIDO_MOV.VCHA_MOV_MOVIMIENTO_ID} = '" + var_clave_movimiento + "'"
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Reporte de Movimientos"
            frmvistasprevias.Show
            Set reporte = Nothing
   
End Sub



Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2900
   Left = 3850
   txt_fecha_inicio = Date
   txt_fecha_fin = Date
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If var_despliega_menu = True Then
      var_swpassword = False
      var_modifica_registro = False
   End If
   Call activa_forma(var_activa_forma_reporte_nivel_surtido)
End Sub


Private Sub mes_LostFocus()
   mes.Visible = False
End Sub

Private Sub txt_fecha_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmcalendario.Show 1
      txt_fecha_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmcalendario.Show 1
      txt_fecha_inicio = var_fecha_general
   End If
End Sub
