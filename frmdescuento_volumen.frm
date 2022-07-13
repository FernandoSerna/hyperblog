VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdescuentos_volumen_grupo_real 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos por volumen a grupo real"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9180
   Begin VB.Frame frm_periodo 
      Height          =   675
      Left            =   3855
      TabIndex        =   26
      Top             =   510
      Width           =   4770
      Begin VB.CommandButton cmd_cambiar_periodo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4350
         Picture         =   "frmdescuento_volumen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Aplicar Pagos Alt + A"
         Top             =   225
         Width           =   330
      End
      Begin VB.ListBox lst_años 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmdescuento_volumen.frx":014A
         Left            =   3405
         List            =   "frmdescuento_volumen.frx":018D
         TabIndex        =   28
         Top             =   255
         Width           =   900
      End
      Begin VB.ComboBox cmb_meses 
         Height          =   315
         ItemData        =   "frmdescuento_volumen.frx":020F
         Left            =   630
         List            =   "frmdescuento_volumen.frx":0237
         TabIndex        =   27
         Top             =   240
         Width           =   2280
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   3015
         TabIndex        =   31
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Left            =   195
         TabIndex        =   30
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.CommandButton cmd_aplicar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmdescuento_volumen.frx":02A0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aplicar Pagos Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmdescuento_volumen.frx":03EA
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Grupos "
      Height          =   2895
      Left            =   105
      TabIndex        =   22
      Top             =   3690
      Width           =   8955
      Begin MSComctlLib.ListView lv_grupos 
         Height          =   2625
         Left            =   60
         TabIndex        =   23
         Top             =   195
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   4630
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cobranza     "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "% Aplicado   "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "% Aplicar   "
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Causa"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   105
      TabIndex        =   20
      Top             =   510
      Width           =   8940
      Begin VB.CommandButton cmd_periodo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8535
         Picture         =   "frmdescuento_volumen.frx":057C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Seleccionar Periodo "
         Top             =   240
         Width           =   330
      End
      Begin VB.TextBox txt_periodo 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   195
         TabIndex        =   21
         Top             =   255
         Width           =   8310
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8670
      Picture         =   "frmdescuento_volumen.frx":06B2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmdescuento_volumen.frx":0CEC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Buscar Movimiento Alt + B"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmdescuento_volumen.frx":0DEE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   90
      Left            =   75
      TabIndex        =   3
      Top             =   315
      Width           =   9015
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Generales "
      Height          =   2370
      Left            =   105
      TabIndex        =   0
      Top             =   1260
      Width           =   8955
      Begin VB.TextBox txt_causa 
         Height          =   315
         Left            =   1740
         TabIndex        =   19
         Top             =   1950
         Width           =   7095
      End
      Begin VB.TextBox txt_descuento_aplicar 
         Height          =   300
         Left            =   1740
         TabIndex        =   12
         Top             =   1620
         Width           =   1380
      End
      Begin VB.TextBox txt_descuento_aplicado 
         Height          =   300
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1290
         Width           =   1380
      End
      Begin VB.TextBox txt_cobranza 
         Height          =   300
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_grupo 
         Height          =   315
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   5700
      End
      Begin VB.TextBox txt_grupo 
         Height          =   300
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   637
         Width           =   1380
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   315
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   5700
      End
      Begin VB.TextBox txt_agente 
         Height          =   300
         Left            =   1740
         TabIndex        =   6
         Top             =   307
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Causa de correción:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2010
         Width           =   1425
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descuento a Aplicar:"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descuento Aplicado:"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   1350
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cobranza:"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   1013
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Real:"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmdescuentos_volumen_grupo_real"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmb_meses_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      lst_años.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_periodo.Visible = False
   End If
End Sub

Private Sub cmd_aplicar_Click()
   Dim si As Integer
   si = MsgBox("¿Deseas aplicar los cambios?", vbYesNo, "ATENCION")
   If si = 6 Then
      rs.Open "update tb_descuentos_volumen_asignado set floa_dvo_descuento_corregido = " + txt_descuento_aplicar + ", vcha_dvo_causa_correccion = '" + txt_causa + "' where vcha_gre_grupo_real_id = '" + txt_grupo + "' and DTIM_DVO_PERIODO_INICIO  = '" + CStr(fecha_inicio) + "' and DTIM_DVO_PERIODO_FIN = '" + CStr(fecha_fin + 1) + "'", cnn, adOpenDynamic
      lv_grupos.selectedItem.SubItems(4) = txt_descuento_aplicar
      lv_grupos.selectedItem.SubItems(5) = txt_causa
   End If
End Sub

Private Sub cmd_cambiar_periodo_Click()
   Dim fecha_anterior As Date
   Dim dia_anterior As Integer
   Dim mes_anterior As Integer
   Dim año_anterior As Integer
   Dim dia As Integer
   Dim mes As Integer
   Dim año As Integer
   Dim periodo As String
   
   If cmb_meses = "Enero" Then
      mes_anterior = 1
   End If
   If cmb_meses = "Febrero" Then
      mes_anterior = 2
   End If
   If cmb_meses = "Marzo" Then
      mes_anterior = 3
   End If
   If cmb_meses = "Abril" Then
      mes_anterior = 4
   End If
   If cmb_meses = "Mayo" Then
      mes_anterior = 5
   End If
   If cmb_meses = "Junio" Then
      mes_anterior = 6
   End If
   If cmb_meses = "Julio" Then
      mes_anterior = 7
   End If
   If cmb_meses = "Agosto" Then
      mes_anterior = 8
   End If
   If cmb_meses = "Septiembre" Then
      mes_anterior = 9
   End If
   If cmb_meses = "Octubre" Then
      mes_anterior = 10
   End If
   If cmb_meses = "Noviembre" Then
      mes_anterior = 11
   End If
   If cmb_meses = "Diciembre" Then
      mes_anterior = 12
   End If
   año_anterior = lst_años
   If mes_anterior = 1 Or mes_anterior = 3 Or mes_anterior = 5 Or mes_anterior = 7 Or mes_anterior = 8 Or mes_anterior = 10 Or mes_anterior = 12 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("31/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 31
   End If
   If mes_anterior = 4 Or mes_anterior = 6 Or mes_anterior = 9 Or mes_anterior = 11 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("30/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 30
   End If
   
   If mes_anterior = 2 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      If año_anterior = 2004 Or año_anterior = 2008 Or año_anterior = 2012 Or año_anterior = 2016 Or año_anterior = 2020 Or año_anterior = 2024 Then
         fecha_fin = CDate("29/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 29
      Else
         fecha_fin = CDate("28/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 28
      End If
   End If
   
   mes = mes_anterior
   año = año_anterior
  
   If mes = 1 Then
      periodo = "Enero"
   End If
   If mes = 2 Then
      periodo = "Febrero"
   End If
   If mes = 3 Then
      periodo = "Marzo"
   End If
   If mes = 4 Then
      periodo = "Abril"
   End If
   If mes = 5 Then
      periodo = "Mayo"
   End If
   If mes = 6 Then
      periodo = "Junio"
   End If
   If mes = 7 Then
      periodo = "Julio"
   End If
   If mes = 8 Then
      periodo = "Agosto"
   End If
   If mes = 9 Then
      periodo = "Septiembre"
   End If
   If mes = 10 Then
      periodo = "Octubre"
   End If
   If mes = 11 Then
      periodo = "Noviembre"
   End If
   If mes = 12 Then
      periodo = "Diciembre"
   End If
   txt_periodo = "1 de " + periodo + " al " + Str(dia_anterior) + " de " + periodo + " del " + Str(año)
   frm_periodo.Visible = False

End Sub

Private Sub cmd_imprimir_Click()
   Set reporte = appl.OpenReport(App.Path + "\rep_descuentos_volumen.rpt")
   reporte.RecordSelectionFormula = "{VW_DESCUENTOS_VOLUMEN.DTIM_DVO_PERIODO_INICIO} = date('" + CStr(fecha_inicio) + "') and {VW_DESCUENTOS_VOLUMEN.DTIM_DVO_PERIODO_FIN} = date('" + CStr(fecha_fin + 1) + "')"
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de Movimientos"
   frmvistasprevias.Show
   Set reporte = Nothing
End Sub

Private Sub cmd_periodo_Click()
   Dim mes As Integer
   Dim año As Integer
   mes = Month(Date)
   año = Year(Date)
   lst_años = año
   If mes = 1 Then
      cmb_meses = "Enero"
   End If
   If mes = 2 Then
      cmb_meses = "Febrero"
   End If
   If mes = 3 Then
      cmb_meses = "Marzo"
   End If
   If mes = 4 Then
      cmb_meses = "Abril"
   End If
   If mes = 5 Then
      cmb_meses = "Mayo"
   End If
   If mes = 6 Then
      cmb_meses = "Junio"
   End If
   If mes = 7 Then
      cmb_meses = "Julio"
   End If
   If mes = 8 Then
      cmb_meses = "Agosto"
   End If
   If mes = 9 Then
      cmb_meses = "Septiembre"
   End If
   If mes = 10 Then
      cmb_meses = "Octubre"
   End If
   If mes = 11 Then
      cmb_meses = "Noviembre"
   End If
   If mes = 12 Then
      cmb_meses = "Diciembre"
   End If
   frm_periodo.Visible = True
   cmb_meses.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
Dim si As Integer
   si = MsgBox("¿Deseas aplicar los descuentos?", vbYesNo, "ATENCION")
   If si = 6 Then
      si = MsgBox("Confirmar la aplicacion de los descuentos", vbYesNo, "ATENCION")
      If si = 6 Then
         rsaux2.Open "update tb_gruposreales set floa_gre_descuento_1 = 0", cnn, adOpenDynamic, adLockOptimistic
         rs.Open "select * from vw_descuentos_volumen", cnn, adOpenDynamic, adLockOptimistic
         
         While Not rs.EOF
            rsaux2.Open "update tb_gruposreales set floa_gre_descuento_1 = " + CStr(rs!floa_dvo_descuento_corregido) + " where vcha_gre_grupo_real_id = '" + rs!vcha_gre_grupo_real_id + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.MoveNext
         Wend
         rs.Close
         MsgBox "Se a terminado de actualizar los descuentos", vbOKOnly, "ATENCION"
      End If
   End If
   
End Sub

Private Sub Form_Load()
   Dim fecha_anterior As Date
   Dim dia_anterior As Integer
   Dim mes_anterior As Integer
   Dim año_anterior As Integer
   Dim dia As Integer
   Dim mes As Integer
   Dim año As Integer
   Dim periodo As String
   frm_periodo.Visible = False
   Top = 200
   Left = 1300
   var_cadena_seguridad = ""
   mes = Month(Date)
   año = Year(Date)
   If mes = 1 Then
      mes_anterior = 12
      año_anterior = año - 1
   Else
      mes_anterior = mes - 1
      año_anterior = año
   End If
   If mes_anterior = 1 Or mes_anterior = 3 Or mes_anterior = 5 Or mes_anterior = 7 Or mes_anterior = 8 Or mes_anterior = 10 Or mes_anterior = 12 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("31/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 31
   End If
   If mes_anterior = 4 Or mes_anterior = 6 Or mes_anterior = 9 Or mes_anterior = 11 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      fecha_fin = CDate("30/" + Str(mes_anterior) + "/" + Str(año_anterior))
      dia_anterior = 30
   End If
   
   If mes_anterior = 2 Then
      fecha_inicio = CDate("1/" + Str(mes_anterior) + "/" + Str(año_anterior))
      If año_anterior = 2004 Or año_anterior = 2008 Or año_anterior = 2012 Or año_anterior = 2016 Or año_anterior = 2020 Or año_anterior = 2024 Then
         fecha_fin = CDate("29/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 29
      Else
         fecha_fin = CDate("28/" + Str(mes_anterior) + "/" + Str(año_anterior))
         dia_anterior = 28
      End If
   End If
   
   mes = mes_anterior
   año = año_anterior
  
   If mes = 1 Then
      periodo = "Enero"
   End If
   If mes = 2 Then
      periodo = "Febrero"
   End If
   If mes = 3 Then
      periodo = "Marzo"
   End If
   If mes = 4 Then
      periodo = "Abril"
   End If
   If mes = 5 Then
      periodo = "Mayo"
   End If
   If mes = 6 Then
      periodo = "Junio"
   End If
   If mes = 7 Then
      periodo = "Julio"
   End If
   If mes = 8 Then
      periodo = "Agosto"
   End If
   If mes = 9 Then
      periodo = "Septiembre"
   End If
   If mes = 10 Then
      periodo = "Octubre"
   End If
   If mes = 11 Then
      periodo = "Noviembre"
   End If
   If mes = 12 Then
      periodo = "Diciembre"
   End If
   txt_periodo = "1 de " + periodo + " al " + Str(dia_anterior) + " de " + periodo + " del " + Str(año)
End Sub


Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_descuentos_volumen_grupo_real)
End Sub

Private Sub lst_años_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmd_cambiar_periodo.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_periodo.Visible = False
   End If
End Sub

Private Sub lv_grupos_GotFocus()
   On Error GoTo salir:
   txt_grupo = lv_grupos.selectedItem
   txt_nombre_grupo = lv_grupos.selectedItem.SubItems(1)
   txt_cobranza = lv_grupos.selectedItem.SubItems(2)
   txt_descuento_aplicado = lv_grupos.selectedItem.SubItems(3)
   txt_descuento_aplicar = lv_grupos.selectedItem.SubItems(4)
   txt_causa = lv_grupos.selectedItem.SubItems(5)
salir:
End Sub

Private Sub lv_grupos_ItemClick(ByVal Item As MSComctlLib.ListItem)
   On Error GoTo salir:
   txt_grupo = lv_grupos.selectedItem
   txt_nombre_grupo = lv_grupos.selectedItem.SubItems(1)
   txt_cobranza = lv_grupos.selectedItem.SubItems(2)
   txt_descuento_aplicado = lv_grupos.selectedItem.SubItems(3)
   txt_descuento_aplicar = lv_grupos.selectedItem.SubItems(4)
   txt_causa = lv_grupos.selectedItem.SubItems(5)
salir:
End Sub

Private Sub txt_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      lv_grupos.SetFocus
   End If
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(txt_agente) <> "" Then
      Dim contador As Integer
      txt_grupo = ""
      txt_nombre_grupo = ""
      txt_cobranza = ""
      txt_descuento_aplicado = ""
      txt_descuento_aplicar = ""
      txt_causa = ""
      rs.Open "select * from tb_Agentes where vcha_age_agente_id = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
         rsaux2.Open "select distinct vcha_gre_grupo_real_id,vcha_gre_nombre, floa_dvo_importe_grupo_real,floa_dvo_descuento, floa_dvo_descuento_corregido, vcha_dvo_causa_correccion from vw_descuentos_volumen where vcha_age_agente_id = '" + txt_agente + "' and DTIM_DVO_PERIODO_INICIO  = '" + CStr(fecha_inicio) + "' and DTIM_DVO_PERIODO_FIN = '" + CStr(fecha_fin + 1) + "'", cnn, adOpenDynamic
         If Not rsaux2.EOF Then
            Dim list_item As ListItem
            contador = 0
            lv_grupos.ListItems.Clear
            While Not rsaux2.EOF
               Set list_item = lv_grupos.ListItems.Add(, , rsaux2!vcha_gre_grupo_real_id)
               list_item.SubItems(1) = IIf(IsNull(rsaux2!vcha_gre_nombre), "", rsaux2!vcha_gre_nombre)
               list_item.SubItems(2) = Format(IIf(IsNull(rsaux2!floa_dvo_importe_grupo_real), 0, rsaux2!floa_dvo_importe_grupo_real), "###,###,##0.00")
               list_item.SubItems(3) = IIf(IsNull(rsaux2!floa_dvo_descuento), 0, rsaux2!floa_dvo_descuento)
               list_item.SubItems(4) = IIf(IsNull(rsaux2!floa_dvo_descuento_corregido), 0, rsaux2!floa_dvo_descuento_corregido)
               list_item.SubItems(5) = IIf(IsNull(rsaux2!vcha_dvo_causa_correccion), "", rsaux2!vcha_dvo_causa_correccion)
               contador = contador + 1
               rsaux2.MoveNext
            Wend
         Else
            lv_grupos.ListItems.Clear
            MsgBox "No existe el calculo del descuento para el periodo seleccionado", vbOKOnly, "ATENCION"
         End If
         rsaux2.Close
         If contador > 11 Then
            lv_grupos.ColumnHeaders(2).Width = 3910
         Else
            lv_grupos.ColumnHeaders(2).Width = 3710
         End If
      End If
      rs.Close
   End If
End Sub

Private Sub txt_descuento_aplicar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txt_causa.SetFocus
   End If
End Sub
