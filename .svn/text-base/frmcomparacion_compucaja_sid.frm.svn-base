VERSION 5.00
Begin VB.Form frmcomparacion_compucaja_sid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comparación   COMPUCAJA - S.I.D."
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   2040
      TabIndex        =   16
      Top             =   -30
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5100
      Picture         =   "frmcomparacion_compucaja_sid.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmcomparacion_compucaja_sid.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   15
      TabIndex        =   12
      Top             =   360
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Caption         =   " Fecha "
      Height          =   645
      Left            =   105
      TabIndex        =   10
      Top             =   465
      Width           =   5400
      Begin VB.TextBox txt_fecha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1710
         TabIndex        =   2
         Top             =   165
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   1140
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Ventas "
      Height          =   1605
      Left            =   105
      TabIndex        =   8
      Top             =   1230
      Width           =   5400
      Begin VB.TextBox txt_venta_compucaja_piezas 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   1170
         TabIndex        =   5
         Top             =   1065
         Width           =   1800
      End
      Begin VB.TextBox txt_venta_compucaja 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   1170
         TabIndex        =   3
         Top             =   585
         Width           =   1800
      End
      Begin VB.TextBox txt_venta_sid_piezas 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   3015
         TabIndex        =   6
         Top             =   1065
         Width           =   1800
      End
      Begin VB.CommandButton cmd_ventas 
         Height          =   330
         Left            =   4905
         Picture         =   "frmcomparacion_compucaja_sid.frx":073C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   870
         Width           =   360
      End
      Begin VB.TextBox txt_ventas_sid 
         Alignment       =   1  'Right Justify
         Height          =   420
         Left            =   3015
         TabIndex        =   4
         Top             =   585
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Compucaja:"
         Height          =   195
         Left            =   1650
         TabIndex        =   15
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Piezas:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1178
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   698
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "S.I.D.:"
         Height          =   195
         Left            =   3690
         TabIndex        =   9
         Top             =   315
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmcomparacion_compucaja_sid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_ventas_Click()
   If IsDate(Me.txt_fecha) Then
      var_dia = CStr(Day(CDate(Me.txt_fecha)))
      var_mes = CStr(Month(CDate(Me.txt_fecha)))
      var_año = CStr(Year(CDate(Me.txt_fecha)))
      If Len(Trim(var_dia)) = 1 Then
         var_dia = "0" + var_dia
      End If
      If Len(Trim(var_mes)) = 1 Then
         var_mes = "0" + var_mes
      End If
      VAR_FECHA_STR = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
      rs.Open "SELECT SUM(CANTIDAD), SUM(PRECIOVENTA) FROM kardex_para_sid where FECHA >= " + VAR_FECHA_STR + " and fecha <" + VAR_FECHA_STR + "+1 and tma_codigo = 2", cnn_compucaja, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_venta_compucaja = Format(IIf(IsNull(rs(1).Value), 0, rs(1).Value), "###,###,##0.00")
         Me.txt_venta_compucaja_piezas = Format(IIf(IsNull(rs(0).Value), 0, rs(0).Value), "###,###,##0.00")
      Else
         Me.txt_venta_compucaja = "0.00"
         Me.txt_venta_compucaja_piezas = "0.00"
      End If
      rs.Close
      rs.Open "select sum(floa_Sal_Cantidad), sum(flOA_sal_Cantidad * floa_sal_precio * 1.16) from tb_Salidas where vcha_mov_movimiento_id = 'CC_2' and dtim_Sal_Fecha >= " + VAR_FECHA_STR + " and dtim_Sal_fecha < " + VAR_FECHA_STR + " +1", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_ventas_sid = Format(IIf(IsNull(rs(1).Value), 0, rs(1).Value), "###,###,##0.00")
         Me.txt_venta_sid_piezas = Format(IIf(IsNull(rs(0).Value), 0, rs(0).Value), "###,###,##0.00")
      Else
         Me.txt_ventas_sid = "0.00"
         Me.txt_venta_sid_piezas = "0.00"
      End If
      rs.Close
   Else
      MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Command1_Click()
   rs.Open "select folio, fecha  from kardex_para_sid where fecha >= {d '2010-07-05'} and fecha < {d '2010-07-06'}"

End Sub

Private Sub Form_Load()
   Top = 2000
   Left = 3200
   Me.txt_fecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_fecha_Change()
   Me.txt_venta_compucaja = ""
   Me.txt_venta_compucaja_piezas = ""
   Me.txt_venta_sid_piezas = ""
   Me.txt_ventas_sid = ""
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      Me.txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_ventas.SetFocus
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_venta_compucaja_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_venta_compucaja_piezas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_venta_sid_piezas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_ventas_sid_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      If KeyAscii = 27 Then
         Unload Me
      Else
         KeyAscii = 0
      End If
   End If
End Sub
