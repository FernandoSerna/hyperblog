VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmenvio_estado_cuenta_correo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio de estado de cuenta por correo"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   75
      Left            =   5880
      TabIndex        =   23
      Top             =   1065
      Width           =   5640
   End
   Begin VB.CommandButton cmd_correo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmenvio_estado_cuenta_correo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Enviar estado de cuenta por correo"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11205
      Picture         =   "frmenvio_estado_cuenta_correo.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmenvio_estado_cuenta_correo.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   60
      TabIndex        =   22
      Top             =   390
      Width           =   11520
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   780
      Left            =   90
      TabIndex        =   19
      Top             =   495
      Width           =   5685
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1005
         TabIndex        =   3
         Top             =   293
         Width           =   1305
      End
      Begin VB.TextBox txt_fin 
         Height          =   330
         Left            =   3525
         TabIndex        =   4
         Top             =   285
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3210
         TabIndex        =   21
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Agentes "
      Height          =   5895
      Left            =   105
      TabIndex        =   18
      Top             =   1335
      Width           =   5670
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1395
         Picture         =   "frmenvio_estado_cuenta_correo.frx":09BE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   735
         Picture         =   "frmenvio_estado_cuenta_correo.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar (Enter)"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Picture         =   "frmenvio_estado_cuenta_correo.frx":0E1E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   75
         Picture         =   "frmenvio_estado_cuenta_correo.frx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   405
         Picture         =   "frmenvio_estado_cuenta_correo.frx":0FF2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   210
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   75
         Left            =   30
         TabIndex        =   24
         Top             =   555
         Width           =   5610
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   5055
         Left            =   90
         TabIndex        =   10
         Top             =   690
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   8916
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Clientes  "
      Height          =   6750
      Left            =   5850
      TabIndex        =   17
      Top             =   480
      Width           =   5700
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1395
         Picture         =   "frmenvio_estado_cuenta_correo.frx":1208
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   735
         Picture         =   "frmenvio_estado_cuenta_correo.frx":141E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Marcar (Enter)"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Picture         =   "frmenvio_estado_cuenta_correo.frx":1668
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   75
         Picture         =   "frmenvio_estado_cuenta_correo.frx":173A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   225
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   405
         Picture         =   "frmenvio_estado_cuenta_correo.frx":183C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   225
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   5835
         Left            =   90
         TabIndex        =   16
         Top             =   765
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   10292
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
      End
   End
End
Attribute VB_Name = "frmenvio_estado_cuenta_correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_fecha_inicio_Change()

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Left = 0
   Top = 0
   rs.Open "select * from tb_agentes where vcha_emp_empresa_id = '" + var_empresa + "' or vcha_age_agente_id = '00083' or vcha_age_Agente_id = '00100'  order by vcha_age_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_ALMACENES = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
      list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      list_item.SubItems(2) = ""
      rs.MoveNext:
   Wend
   rs.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = CDate(Date)
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lv_agentes.SetFocus
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = CDate(Date)
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_fin.SetFocus
   End If
End Sub
