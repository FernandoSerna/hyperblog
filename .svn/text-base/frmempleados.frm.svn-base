VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#11.0#0"; "vbskfree.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmempleados 
   Caption         =   "Empleados"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   3240
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin MSComCtl2.MonthView m1 
      Height          =   2370
      Left            =   2640
      TabIndex        =   46
      Top             =   3720
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   53870594
      CurrentDate     =   37440
   End
   Begin VB.Frame empleados 
      Height          =   6015
      Left            =   120
      TabIndex        =   44
      Top             =   0
      Width           =   2175
      Begin MSFlexGridLib.MSFlexGrid cuadro 
         Height          =   5850
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   10319
         _Version        =   393216
         Cols            =   7
         BackColor       =   16777215
         BackColorFixed  =   -2147483637
         GridColorFixed  =   16777215
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.Frame datos3 
      Height          =   1695
      Left            =   2400
      TabIndex        =   38
      Top             =   4320
      Width           =   8415
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   7080
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   4800
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   4800
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox operaciones 
         Height          =   315
         ItemData        =   "frmempleados.frx":0000
         Left            =   1200
         List            =   "frmempleados.frx":000A
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factor"
         Height          =   195
         Index           =   18
         Left            =   6600
         TabIndex        =   43
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promedio Destajo"
         Height          =   195
         Index           =   17
         Left            =   3000
         TabIndex        =   42
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operacion"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sueldo Diario Integrado"
         Height          =   195
         Index           =   15
         Left            =   3000
         TabIndex        =   40
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sueldo Diario"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Frame datos2 
      Height          =   1215
      Left            =   2400
      TabIndex        =   32
      Top             =   3000
      Width           =   8415
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   4800
         TabIndex        =   45
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton c2 
         Height          =   255
         Left            =   6480
         Picture         =   "frmempleados.frx":0028
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton c1 
         Height          =   255
         Left            =   2880
         Picture         =   "frmempleados.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   255
      End
      Begin VB.ComboBox nacionalidad 
         Height          =   315
         ItemData        =   "frmempleados.frx":02BC
         Left            =   4800
         List            =   "frmempleados.frx":02C6
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   34
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Baja"
         Height          =   195
         Index           =   13
         Left            =   3600
         TabIndex        =   37
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alta"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad"
         Height          =   195
         Index           =   11
         Left            =   3600
         TabIndex        =   35
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. IMSS"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame datos1 
      Height          =   2175
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      Begin VB.ComboBox sexo 
         Height          =   315
         ItemData        =   "frmempleados.frx":02E0
         Left            =   960
         List            =   "frmempleados.frx":02EA
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   4800
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   3240
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox poblacion 
         Height          =   315
         ItemData        =   "frmempleados.frx":0303
         Left            =   960
         List            =   "frmempleados.frx":030A
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox civil 
         Height          =   315
         ItemData        =   "frmempleados.frx":031E
         Left            =   3240
         List            =   "frmempleados.frx":032E
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   4800
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   5640
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
         Height          =   195
         Index           =   9
         Left            =   2280
         TabIndex        =   31
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colonia"
         Height          =   195
         Index           =   7
         Left            =   4200
         TabIndex        =   29
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R.F.C."
         Height          =   195
         Index           =   6
         Left            =   4200
         TabIndex        =   28
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.P."
         Height          =   195
         Index           =   5
         Left            =   2880
         TabIndex        =   27
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Poblacion"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos"
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   24
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   405
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":035A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":08F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":0E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":1428
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":19C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":1F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":2A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":302A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":35C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":3B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmempleados.frx":40F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   360
      Left            =   2400
      TabIndex        =   21
      Top             =   240
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmempleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub c1_Click()
    m1.Visible = True
    m1.SetFocus
End Sub

Private Sub c2_Click()
    m1.Visible = True
    m1.SetFocus
End Sub

Private Sub c2_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cuadro_Click()
    'Call pro_consulta_empleado(Me, "tb_empleados", "em_nombre", Trim(cuadro.Text))
End Sub

Private Sub cuadro_EnterCell()
   ' Call pro_consulta_empleado(Me, "tb_empleados", "em_nombre", Trim(cuadro.Text))
End Sub

Private Sub cuadro_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Call SortByColumn(Me, cuadro.MouseCol)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call menuvisible(Frmmenu2, True)
End Sub

Private Sub m1_KeyPress(KeyAscii As Integer)
    Me.KeyPreview = False
    If KeyAscii = 13 Then
        Text1(9) = m1.Value
        m1.Visible = False
        c2.SetFocus
    End If
End Sub

Private Sub sexo_LostFocus()
    Call pro_combodrop(civil, True)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Select Case Index
    Case 5: Call pro_combodrop(poblacion, True)
    Case 7: Call pro_combodrop(sexo, True)
    Case 8: Call pro_combodrop(nacionalidad, True)
    Case 12: Call pro_combodrop(operaciones, True)
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index

    Case 2
     '   Call pro_guardar_empleado(Me)
     '   Call llena_grids(Me, "tb_empleados", 1)
        
    End Select
        Call pro_limpiatextos(Me)
        Text1(0).SetFocus
End Sub
