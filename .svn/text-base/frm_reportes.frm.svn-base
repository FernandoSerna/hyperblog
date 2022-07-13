VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "VBSKFREE.OCX"
Begin VB.Form frmreportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTES"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frm_reportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_inventario_historico 
      Caption         =   "Generando Inventario Historico"
      Height          =   615
      Left            =   600
      TabIndex        =   37
      Top             =   2880
      Visible         =   0   'False
      Width           =   4935
      Begin MSComctlLib.ProgressBar pb_1 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComCtl2.MonthView mon_fecha_2 
      Height          =   2370
      Left            =   3240
      TabIndex        =   34
      Top             =   4320
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   19726337
      CurrentDate     =   37484
   End
   Begin MSComCtl2.MonthView mon_fecha_1 
      Height          =   2370
      Left            =   600
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   19726337
      CurrentDate     =   37484
   End
   Begin VB.Frame fra_codigo 
      Caption         =   "Codigo"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   6135
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame fra_ejecuta 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   6135
      Begin VB.OptionButton Opt_pantalla 
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Opt_impresora 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Opt_excel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Opt_correo 
         Caption         =   "Correo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   600
         Picture         =   "frm_reportes.frx":08CA
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2040
         Picture         =   "frm_reportes.frx":1194
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   3480
         Picture         =   "frm_reportes.frx":1A5E
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   5040
         Picture         =   "frm_reportes.frx":2328
         Top             =   600
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   1440
         Y1              =   360
         Y2              =   1200
      End
      Begin VB.Line Line3 
         X1              =   1800
         X2              =   3120
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line4 
         X1              =   3120
         X2              =   3120
         Y1              =   360
         Y2              =   1200
      End
      Begin VB.Line Line5 
         X1              =   3360
         X2              =   4560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line6 
         X1              =   4560
         X2              =   4560
         Y1              =   360
         Y2              =   1200
      End
      Begin VB.Line Line7 
         X1              =   4800
         X2              =   6000
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line8 
         X1              =   6000
         X2              =   6000
         Y1              =   360
         Y2              =   1200
      End
   End
   Begin VB.Frame fra_fechas 
      Caption         =   "Fechas"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   6135
      Begin MSMask.MaskEdBox mas_fecha_inicial 
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mas_fecha_final 
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Final"
         Height          =   195
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicial"
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame fra_descripcion 
      Caption         =   "Descripcion"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   5880
      Width           =   6135
      Begin VB.TextBox txt_descripcion 
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame fra_sublineas 
      Caption         =   "Sub Lineas"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   6135
      Begin VB.ComboBox cbo_sublinea_1 
         Height          =   315
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cbo_sublinea_2 
         Height          =   315
         Left            =   4440
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Sub Linea Inicial"
         Height          =   195
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Sub Linea Final"
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Frame fra_lineas 
      Caption         =   "Lineas"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   6135
      Begin VB.ComboBox cbo_linea_2 
         Height          =   315
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cbo_linea_1 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Linea Final"
         Height          =   195
         Left            =   3480
         TabIndex        =   15
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Linea Inicial"
         Height          =   195
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Incluir Articulos Sin Existencia"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   5520
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.Frame fra_fecha 
      Caption         =   "Aque Fecha Quiere el Reporte"
      Height          =   2655
      Left            =   3240
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
      Begin MSComCtl2.MonthView mon_fecha_inventario 
         Height          =   2370
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   19726337
         CurrentDate     =   37491
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reportes del Almacen"
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6135
      Begin VB.ListBox lis_reportes 
         Height          =   1425
         ItemData        =   "frm_reportes.frx":2BF2
         Left            =   120
         List            =   "frm_reportes.frx":2C02
         TabIndex        =   7
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame fra_opciones 
      Caption         =   "Opciones de Filtro"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   6135
      Begin VB.CheckBox Check6 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Sublineas"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Lineas"
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Fechas"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmreportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cadena As String
Dim fecha1 As Date, fecha2 As Date, var_fecha_final As Date


Sub pro_imprime()

    'CR1.CopiesToPrinter = 1
    Select Case lis_reportes.ListIndex
    Case 0
        CR1.SelectionFormula = Cadena
        CR1.ReportFileName = App.Path + "\rep_inventario.rpt"
    Case 1
        CR1.SelectionFormula = Cadena
        CR1.ReportFileName = App.Path + "\rep_inventario_costo.rpt"
    Case 2
        CR1.SelectionFormula = Cadena
        CR1.ReportFileName = App.Path + "\rep_movimientos.rpt"
       ' CR1.GroupSortFields(0) = "{TB_DETALLE.BINT_DET_FOLIO}"
        CR1.Formulas(3) = "fecha1= """ & mas_fecha_inicial & """"
        CR1.Formulas(4) = "fecha2= """ & mas_fecha_final & """"
        If Check3.Value = 1 Then
            CR1.Formulas(5) = "linea= """ & Check3.caption & """"
        End If
        If Check4.Value = 1 Then
            CR1.Formulas(6) = "sublinea= """ & Check4.caption & """"
        End If
        If Check6.Value = 1 Then
            CR1.Formulas(7) = "codigo= """ & Check5.caption & """"
        End If
    Case 3
        CR1.SelectionFormula = Cadena
        CR1.ReportFileName = App.Path + "\rep_inventario_h1.rpt"
        CR1.Formulas(3) = "fecha= """ & mas_fecha_inicial & """"
    End Select
   
    CR1.Formulas(0) = "usuario= """ & fun_NombreUsuario & """"
    CR1.Formulas(1) = "maquina= """ & fun_NombrePc & """"
    CR1.Formulas(2) = "hora= """ & Time & """"
    CR1.RetrieveDataFiles
    CR1.Action = 1
    Cadena = ""
End Sub


Private Sub cbo_linea_1_KeyPress(KeyAscii As Integer)
    
    Call pro_valida_enter(KeyAscii)
    If KeyAscii = 13 Then
        Call pro_combodrop(cbo_linea_2, True)
        cbo_linea_2.SetFocus
    End If
End Sub

Private Sub cbo_linea_2_KeyPress(KeyAscii As Integer)
    
    Call pro_valida_enter(KeyAscii)

End Sub

Sub pro_valida_enter(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If

End Sub


Private Sub cbo_sublinea_1_KeyPress(KeyAscii As Integer)
    Call pro_valida_enter(KeyAscii)
    If KeyAscii = 13 Then
        Call pro_combodrop(cbo_sublinea_2, True)
        cbo_sublinea_2.SetFocus
    End If

End Sub

Private Sub Check2_Click()
    
    If Check2.Value = 1 Then
        fra_fechas.Enabled = True
        fra_fechas.ForeColor = vbBlue: fra_fechas.FontBold = True
        mas_fecha_inicial.SetFocus
    Else
        fra_fechas.ForeColor = vbWhite: fra_fechas.FontBold = False
        fra_fechas.Enabled = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        fra_lineas.Enabled = True
        fra_lineas.ForeColor = vbBlue: fra_lineas.FontBold = True
        Call pro_combodrop(cbo_linea_1, True)
        cbo_linea_1.SetFocus
    Else
        fra_lineas.ForeColor = vbWhite: fra_lineas.FontBold = False
        fra_lineas.Enabled = False
        cbo_linea_1 = "": cbo_linea_2 = ""
    End If

End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        fra_sublineas.Enabled = True
        fra_sublineas.ForeColor = vbBlue: fra_sublineas.FontBold = True
        Call pro_combodrop(cbo_sublinea_1, True)
        cbo_sublinea_1.SetFocus
    Else
        fra_sublineas.ForeColor = vbWhite: fra_sublineas.FontBold = False
        fra_sublineas.Enabled = False
        cbo_sublinea_1 = "": cbo_sublinea_2 = ""
    End If

End Sub

Private Sub Check5_Click()
    If Check5.Value = 1 Then
        fra_descripcion.Enabled = True
        fra_descripcion.ForeColor = vbBlue: fra_descripcion.FontBold = True
        txt_descripcion.SetFocus
    Else
        fra_descripcion.ForeColor = vbWhite: fra_descripcion.FontBold = False
        fra_descripcion.Enabled = False
        txt_descripcion = ""
    End If

End Sub

Private Sub Check6_Click()
    If Check6.Value = 1 Then
        fra_codigo.Enabled = True
        fra_codigo.ForeColor = vbBlue: fra_codigo.FontBold = True
        txt_codigo.SetFocus
    Else
        fra_codigo.ForeColor = vbWhite: fra_codigo.FontBold = False
        txt_codigo = ""
        fra_codigo.Enabled = False
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(ucas(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    rs.Open "SELECT DISTINCT VCHA_LIN_LINEA FROM TB_lineas_VIEW", cnn, adOpenKeyset, adLockOptimistic, adCmdText
    Call RecsetToCombo(cbo_linea_1.hwnd, rs, 0)
    rs.Close

    rs.Open "SELECT DISTINCT VCHA_LIN_LINEA FROM TB_lineas_VIEW", cnn, adOpenKeyset, adLockOptimistic, adCmdText
    Call RecsetToCombo(cbo_linea_2.hwnd, rs, 0)
    rs.Close

    rs.Open "SELECT DISTINCT VCHA_LIN_SUBLINEA FROM TB_lineas_VIEW WHERE VCHA_LIN_SUBLINEA <>''", cnn, adOpenKeyset, adLockOptimistic, adCmdText
    Call RecsetToCombo(cbo_sublinea_1.hwnd, rs, 0)
    rs.Close

    rs.Open "SELECT DISTINCT VCHA_LIN_SUBLINEA FROM TB_lineas_VIEW WHERE VCHA_LIN_SUBLINEA <>''", cnn, adOpenKeyset, adLockOptimistic, adCmdText
    Call RecsetToCombo(cbo_sublinea_2.hwnd, rs, 0)
    rs.Close
    
    mas_fecha_inicial = Format(Date, "dd/mm/yyyy")
    mas_fecha_final = Format(Date, "dd/mm/yyyy")
    mon_fecha_inventario = Format(Date, "dd/mm/yyyy")
    mon_fecha_1 = Format(Date, "dd/mm/yyyy")
    mon_fecha_2 = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call menuvisible(Frmmenu2, True)

End Sub


Private Sub txt_parametros_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_fecha_final_KeyPress(KeyAscii As Integer)

    Call pro_valida_enter(KeyAscii)
    
End Sub

Private Sub txt_fecha_inicial_KeyPress(KeyAscii As Integer)
    
    Call pro_valida_enter(KeyAscii)
    
End Sub

Private Sub List1_Click()
End Sub

Private Sub lis_reportes_Click()
    fra_opciones.Enabled = True
    fra_ejecuta.Enabled = True
    Check2.Enabled = True
    If lis_reportes.ListIndex = 3 Then
        fra_fecha.Visible = True
        mon_fecha_inventario.SetFocus
    End If
    
End Sub


Private Sub Option1_Click()

End Sub

Private Sub mas_fecha_final_GotFocus()
    mas_fecha_final.BackColor = &HC0FFC0
'    mas_fecha_final.SelStart = 0
'    mas_fecha_final.SelLength = Len(mas_fecha_final.Text)

End Sub

Private Sub mas_fecha_final_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = 0
    Else
        mon_fecha_2.Visible = True
        mon_fecha_2.SetFocus
        mon_fecha_2.Value = Date
    End If
End Sub

Private Sub mas_fecha_final_LostFocus()
    
    mas_fecha_final.BackColor = &H80000005

End Sub


Private Sub mas_fecha_inicial_GotFocus()
    mas_fecha_inicial.BackColor = &HC0FFC0
'    mas_fecha_inicial.SelStart = 0
'    mas_fecha_inicial.SelLength = Len(mas_fecha_inicial.Text)

End Sub

Private Sub mas_fecha_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = 0
    Else
        mon_fecha_1.Visible = True
        mon_fecha_1.SetFocus
        mon_fecha_1.Value = Date
    End If
End Sub



Private Sub mas_fecha_inicial_LostFocus()

    mas_fecha_inicial.BackColor = &H80000005

End Sub

Private Sub mon_fecha_1_DateClick(ByVal DateClicked As Date)
    mas_fecha_inicial = Format(mon_fecha_1.Value, "dd/mm/yyyy")
    mon_fecha_1.Visible = False
    mas_fecha_final.SetFocus
End Sub

Private Sub mon_fecha_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        mon_fecha_1.Visible = False: mas_fecha_inicial.SetFocus
    End If
    If KeyAscii = 13 Then
        mas_fecha_inicial = Format(mon_fecha_1.Value, "dd/mm/yyyy")
        mon_fecha_1.Visible = False
        mas_fecha_final.SetFocus
    End If
End Sub

Private Sub mon_fecha_2_DateClick(ByVal DateClicked As Date)
    mas_fecha_final = Format(mon_fecha_2.Value, "dd/mm/yyyy")
    mon_fecha_2.Visible = False
End Sub

Private Sub mon_fecha_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        mon_fecha_2.Visible = False: mas_fecha_final.SetFocus
    End If
    If KeyAscii = 13 Then
        mas_fecha_final = Format(mon_fecha_2.Value, "dd/mm/yyyy")
        mon_fecha_2.Visible = False
    End If
End Sub

Private Sub mon_fecha_inventario_DateClick(ByVal DateClicked As Date)
    var_fecha_final = Format(mon_fecha_inventario, "dd/mm/yyyy")
    mon_fecha_inventario = Format(Date, "dd/mm/yyyy")
    fra_fecha.Visible = False
    pro_genera_historico
End Sub

Private Sub mon_fecha_inventario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        mon_fecha_inventario = Format(Date, "dd/mm/yyyy")
        fra_fecha.Visible = False
        var_fecha_final = Format(Date, "dd/mm/yyyy")
    End If
    If KeyAscii = 13 Then
        var_fecha_final = Format(mon_fecha_inventario, "dd/mm/yyyy")
        mon_fecha_inventario = Format(Date, "dd/mm/yyyy")
        fra_fecha.Visible = False
        pro_genera_historico
    End If
End Sub

Public Sub pro_genera_historico()
Dim RS1 As adodb.Recordset
Dim var_a_fecha As Date
Dim i As Long, x As Byte
Dim afecta As String
    
    If IsDate(var_fecha_final) Then
        fra_inventario_historico.Visible = True
        Check2.Enabled = False

        x = 0
        var_a_fecha = var_fecha_final
        mes = Month(var_fecha_final)
        mes2 = mes - 1
        var_a_fecha = Replace(var_a_fecha, Format(mes, "00"), Format(mes2, "00"))
        rs.Open "select * from TB_CIERRE where month(TB_CIERRE.VCHA_CIE_FECHA)= '" & Month(var_a_fecha) & "'", cnn, adOpenDynamic, adLockOptimistic
        If rs.RecordCount <> 0 Then
            var_a_fecha = IIf(IsNull(rs(25).Value), "", rs(25).Value)
            rsaux.Open "select * from TB_RESULTADO", cnn, adOpenDynamic, adLockOptimistic
            If rsaux.RecordCount <> 0 Then
                rsaux.Close: rsaux.Open "delete from TB_RESULTADO", cnn, adOpenDynamic, adLockOptimistic
                rsaux.Open "select * from TB_RESULTADO", cnn, adOpenDynamic, adLockOptimistic
            End If
            While Not rs.EOF
                rsaux.AddNew
                rsaux(0).Value = IIf(IsNull(rs(0).Value), "", rs(0).Value)
                rsaux(1).Value = IIf(IsNull(rs(1).Value), "", rs(1).Value)
                rsaux(2).Value = IIf(IsNull(rs(2).Value), "", rs(2).Value)
                rsaux(3).Value = IIf(IsNull(rs(3).Value), "", rs(3).Value)
                rsaux(4).Value = IIf(IsNull(rs(4).Value), "", rs(4).Value)
                rsaux(5).Value = IIf(IsNull(rs(5).Value), "", rs(5).Value)
                rsaux(6).Value = IIf(IsNull(rs(6).Value), "", rs(6).Value)
                rsaux(7).Value = IIf(IsNull(rs(7).Value), "", rs(7).Value)
                rsaux(8).Value = IIf(IsNull(rs(8).Value), "", rs(8).Value)
                rsaux(9).Value = IIf(IsNull(rs(9).Value), "", rs(9).Value)
                rsaux(10).Value = IIf(IsNull(rs(10).Value), 0, rs(10).Value)
                rsaux(11).Value = IIf(IsNull(rs(11).Value), 0, rs(11).Value)
                rsaux(12).Value = IIf(IsNull(rs(12).Value), 0, rs(12).Value)
                rsaux(13).Value = IIf(IsNull(rs(13).Value), 0, rs(13).Value)
                rsaux(14).Value = IIf(IsNull(rs(14).Value), 0, rs(14).Value)
                rsaux(15).Value = IIf(IsNull(rs(15).Value), 0, rs(15).Value)
                rsaux(16).Value = IIf(IsNull(rs(16).Value), "", rs(16).Value)
                rsaux(17).Value = IIf(IsNull(rs(17).Value), "", rs(17).Value)
                rsaux(18).Value = IIf(IsNull(rs(18).Value), "", rs(18).Value)
                rsaux(19).Value = IIf(IsNull(rs(19).Value), "", rs(19).Value)
                rsaux(20).Value = IIf(IsNull(rs(20).Value), "", rs(20).Value)
                rsaux(21).Value = IIf(IsNull(rs(21).Value), "", rs(21).Value)
                rsaux(22).Value = IIf(IsNull(rs(22).Value), "", rs(22).Value)
                rsaux(23).Value = IIf(IsNull(rs(23).Value), "", rs(23).Value)
                rsaux(24).Value = IIf(IsNull(rs(24).Value), "", rs(24).Value)
                rsaux(25).Value = IIf(IsNull(rs(25).Value), "", rs(25).Value)
                rsaux.Update
                rs.MoveNext
                x = x + 1
                If x = 100 Then x = 0
                PB_1.Value = x
            Wend
            PB_1.Value = 100
            rs.Close: rsaux.Close
            rs.Open "select * from TB_DETALLE WHERE TB_DETALLE.DTIM_AUD_FECHA BETWEEN '" & var_a_fecha & "' and '" & var_fecha_final & "'", cnn, adOpenDynamic, adLockOptimistic
            If rs.RecordCount <> 0 Then
                While Not rs.EOF
                    rsaux.Open "select * from TB_RESULTADO WHERE VCHA_RES_RESULTADO_ID = '" & rs(1).Value & "'", cnn, adOpenDynamic, adLockOptimistic
                    If rsaux.RecordCount <> 0 Then
                        If rs(6).Value = "SUMA" Then
                            rsaux(17) = Val(rsaux(17)) + Val(rs(2))
                        Else
                            rsaux(17) = Val(rsaux(17)) - Val(rs(2))
                        End If
                        rsaux.Update
                    End If
                    rs.MoveNext
                    rsaux.Close
                Wend
            End If
            rs.Close
        End If
        fra_inventario_historico.Visible = False
        SetTimer hwnd, NV_CLOSEMSGBOX, 1400, AddressOf TimerProc
        MsgBox "Se Genero Exitosamente a la Fecha " & Str(var_fecha_final), , "TRANSACCIONES [ AVISO ]"
    Else
        SetTimer hwnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
        MsgBox "Fecha Invalida... ", , "TRANSACCIONES [ AVISO ]"
    End If

End Sub
Private Sub Opt_correo_Click()
    
'    CR1.Destination = crptMapi
'    pro_opciones
'    pro_imprime
'    Cadena = ""
'    Opt_pantalla.Value = False
'    Call pro_Vacia_formulas(CR1, 3)

End Sub

Private Sub Opt_excel_Click()
    
    CR1.PrintFileType = crptExcel50
    CR1.Destination = crptToFile
    CR1.PrintFileName = App.Path + "reporte\reporte.xls"
    pro_opciones
    pro_imprime
    Cadena = ""
    Opt_pantalla.Value = False
  '  Call pro_Vacia_formulas(CR1, 3)

End Sub

Private Sub Opt_impresora_Click()
    
    CR1.Destination = crptToPrinter
    pro_opciones
    pro_imprime
    Cadena = ""
    Opt_impresora.Value = False
    'Call pro_Vacia_formulas(CR1, 3)

End Sub

Private Sub Opt_pantalla_Click()
    
    CR1.Destination = crptToWindow
    CR1.WindowState = crptMaximized
    pro_opciones
    pro_imprime
    Cadena = ""
    Opt_pantalla.Value = False
   ' Call pro_Vacia_formulas(CR1, 6)
    Check2.Enabled = True
End Sub

Public Sub pro_opciones()
    If Check2 Then
        fecha1 = CDate(mas_fecha_inicial)
        fecha2 = CDate(mas_fecha_final)
        dia1 = Day(fecha1): Mes1 = Month(fecha1): ano1 = Year(fecha1)
        dia2 = Day(fecha2): mes2 = Month(fecha2): ano2 = Year(fecha2)
        If lis_reportes.ListIndex = 2 Then
            If Cadena <> "" Then
                Cadena = "and {tb_transacciones.DTIM_AUD_FECHA} in date(" & ano1 & "," & Mes1 & "," & dia1 & ") to Date(" & ano2 & "," & mes2 & "," & dia2 & ")" '"
            Else
                Cadena = "{tb_transacciones.DTIM_AUD_FECHA} in date(" & ano1 & "," & Mes1 & "," & dia1 & ") to Date(" & ano2 & "," & mes2 & "," & dia2 & ")"
            End If
        Else
            If Cadena <> "" Then
                Cadena = "and {tb_articulos.DTIM_AUD_FECHA} in date(" & ano1 & "," & Mes1 & "," & dia1 & ") to Date(" & ano2 & "," & mes2 & "," & dia2 & ")" '"
            Else
                Cadena = "{tb_articulos.DTIM_AUD_FECHA} in date(" & ano1 & "," & Mes1 & "," & dia1 & ") to Date(" & ano2 & "," & mes2 & "," & dia2 & ")"
            End If
        End If
    Else
        Cadena = ""
    End If
    If Check3 Then
        If lis_reportes.ListIndex <> 3 Then
            If Cadena <> "" Then
                Cadena = Cadena + "and {tb_articulos.vcha_art_linea} in ('" & cbo_linea_1 & "') to ('" & cbo_linea_2 & "')"
            Else
                Cadena = "{tb_articulos.vcha_art_linea} in ('" & cbo_linea_1 & "') to ('" & cbo_linea_2 & "')"
            End If
        Else
            If Cadena <> "" Then
                Cadena = Cadena + "and {tb_resultado.vcha_art_linea} in ('" & cbo_linea_1 & "') to ('" & cbo_linea_2 & "')"
            Else
                Cadena = "{tb_resultado.vcha_art_linea} in ('" & cbo_linea_1 & "') to ('" & cbo_linea_2 & "')"
            End If
        End If
    Else
        If Cadena = "" Then Cadena = ""
    End If
    If Check4 Then
        If lis_reportes.ListIndex <> 3 Then
            If Cadena <> "" Then
                Cadena = Cadena + "and {tb_articulos.vcha_art_sublinea} in ('" & cbo_sublinea_1 & "') to ('" & cbo_sublinea_2 & "')"
            Else
                Cadena = "{tb_articulos.vcha_art_sublinea} in ('" & cbo_sublinea_1 & "') to ('" & cbo_sublinea_2 & "')"
            End If
        Else
            If Cadena <> "" Then
                Cadena = Cadena + "and {tb_resultado.vcha_art_sublinea} in ('" & cbo_sublinea_1 & "') to ('" & cbo_sublinea_2 & "')"
            Else
                Cadena = "{tb_resultado.vcha_art_sublinea} in ('" & cbo_sublinea_1 & "') to ('" & cbo_sublinea_2 & "')"
            End If
        End If
    Else
        If Cadena = "" Then Cadena = ""
    End If
    If Check5 Then
        If lis_reportes.ListIndex <> 3 Then
            If Cadena <> "" Then
                Cadena = Cadena + "and {tb_articulos.vcha_art_descripcion} = '" & txt_descripcion & "'"
            Else
                Cadena = "{tb_articulos.vcha_art_descripcion} = '" & txt_descripcion & "'"
            End If
        Else
            If Cadena <> "" Then
                Cadena = Cadena + "and {tb_resultado.vcha_art_descripcion} = '" & txt_descripcion & "'"
            Else
                Cadena = "{tb_resultado.vcha_art_descripcion} = '" & txt_descripcion & "'"
            End If
        End If
    Else
        If Cadena = "" Then Cadena = ""
    End If
    If Check6 Then
        If lis_reportes.ListIndex <> 3 Then
            If Cadena <> "" Then
                Cadena = Cadena + "and {tb_articulos.VCHA_ART_ARTICULO_ID} = '" & txt_codigo & "'"
            Else
                Cadena = "{tb_articulos.VCHA_ART_ARTICULO_ID} = '" & txt_codigo & "'"
            End If
        Else
            If Cadena <> "" Then
                Cadena = Cadena + "and {tb_resultado.VCHA_RES_RESULTADO_ID} = '" & txt_codigo & "'"
            Else
                Cadena = "{tb_resultado.VCHA_RES_RESULTADO_ID} = '" & txt_codigo & "'"
            End If
        End If
    Else
        If Cadena = "" Then Cadena = ""
    End If
End Sub

'Public Sub pro_Vacia_formulas(listado As Crystal.CrystalReport, iNumero As Integer)
    
'    Dim tiForm As Integer
    
'    For tiForm = 0 To iNumero
'        listado.Formulas(tiForm) = ""
'    Next tiForm
    
'    For tiForm = 0 To 10
'        listado.SortFields(tiForm) = ""
'    Next tiForm

'End Sub

Private Sub txt_codigo_GotFocus()

    txt_codigo.BackColor = &HC0FFC0
    txt_codigo.SelStart = 0
    txt_codigo.SelLength = Len(txt_codigo.Text)

End Sub

Private Sub txt_codigo_LostFocus()

    txt_codigo.BackColor = &H80000005

End Sub

Private Sub txt_descripcion_GotFocus()

    txt_descripcion.BackColor = &HC0FFC0
    txt_descripcion.SelStart = 0
    txt_descripcion.SelLength = Len(txt_descripcion.Text)

End Sub

Private Sub txt_descripcion_LostFocus()

    txt_descripcion.BackColor = &H80000005
    
End Sub
