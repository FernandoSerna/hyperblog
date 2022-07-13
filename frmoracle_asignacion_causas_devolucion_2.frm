VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_asignacion_causas_devolucion_2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de causas de devolución"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   3000
      Left            =   7920
      TabIndex        =   24
      Top             =   600
      Width           =   5970
      Begin VB.TextBox txt_filtro 
         Height          =   405
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   2550
         Left            =   60
         TabIndex        =   26
         Top             =   405
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   4498
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
         BackColor       =   &H000000C0&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   27
         Top             =   120
         Width           =   5895
      End
   End
   Begin VB.Frame Frame3 
      Height          =   135
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   15135
   End
   Begin VB.TextBox txt_defecto 
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
      Left            =   12480
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txt_usuario 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   4980
      Left            =   15
      TabIndex        =   7
      Top             =   3585
      Width           =   15315
      Begin VB.Frame frm_mensaje 
         Height          =   1215
         Left            =   570
         TabIndex        =   12
         Top             =   3645
         Width           =   7065
         Begin VB.Label lbl_mensaje 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "estatus"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1200
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   7050
         End
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   705
         Picture         =   "frmoracle_asignacion_causas_devolucion_2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar (Enter)"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1035
         Picture         =   "frmoracle_asignacion_causas_devolucion_2.frx":024A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_asignacion_causas_devolucion_2.frx":031C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_asignacion_causas_devolucion_2.frx":041E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_devoluciones 
         Height          =   4755
         Left            =   45
         TabIndex        =   14
         Top             =   120
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   8387
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2478
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Causa original"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Agente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cliente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Causa real"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Folio ANC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Usuario"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   14925
      Picture         =   "frmoracle_asignacion_causas_devolucion_2.frx":0634
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   300
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmoracle_asignacion_causas_devolucion_2.frx":0C6E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.TextBox txt_codigo 
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
      Left            =   4320
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txt_folio_ANC 
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
      Left            =   8520
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   0
      TabIndex        =   6
      Top             =   270
      Width           =   15240
   End
   Begin VB.Label lbl_nombre_cliente 
      Height          =   255
      Left            =   4200
      TabIndex        =   30
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lbl_nombre_agente 
      Height          =   255
      Left            =   2640
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl_defecto_original 
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl_anc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   2280
      Width           =   15015
   End
   Begin VB.Label lbl_defecto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   15015
   End
   Begin VB.Label lbl_articulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   15015
   End
   Begin VB.Label lbl_usuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   15015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Defecto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11400
      TabIndex        =   18
      Top             =   540
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   17
      Top             =   540
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SKU:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   16
      Top             =   540
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Folio ANC:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7200
      TabIndex        =   15
      Top             =   540
      Width           =   1290
   End
End
Attribute VB_Name = "frmoracle_asignacion_causas_devolucion_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim var_orden As String
Dim var_posicion As Double
Dim var_indice As String
Dim var_encontro As Double
Dim var_numero As Double
Dim var_clave_movimiento As String
Dim var_causa_original As String
Dim var_descripcion As String
Dim var_cantidad As Double
Dim var_folio_Devolucion As String
Private Sub cmd_invertir_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
         lv_devoluciones.selectedItem.SubItems(4) = ""
         lv_devoluciones.ListItems.Item(i).Bold = False
         lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
      Else
         lv_devoluciones.selectedItem.SubItems(4) = "*"
         lv_devoluciones.ListItems.Item(i).Bold = True
         lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
      End If
   Next i
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If
End Sub

Private Sub cmd_marcar_Click()
   i = lv_devoluciones.selectedItem.Index
   If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
      lv_devoluciones.selectedItem.SubItems(4) = ""
      lv_devoluciones.ListItems.Item(i).Bold = False
      lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
      lv_devoluciones.Refresh
   Else
      lv_devoluciones.selectedItem.SubItems(4) = "*"
      lv_devoluciones.ListItems.Item(i).Bold = True
      lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
      lv_devoluciones.Refresh
   End If
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If
End Sub

Private Sub cmd_ninguno_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      lv_devoluciones.selectedItem.SubItems(4) = ""
      lv_devoluciones.ListItems.Item(i).Bold = False
      lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = False
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
   Next i
   lv_devoluciones.Refresh
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If

End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_todos_Click()
   n = lv_devoluciones.ListItems.Count
   For i = 1 To n
      lv_devoluciones.ListItems.Item(i).Selected = True
      lv_devoluciones.selectedItem.SubItems(4) = "*"
      lv_devoluciones.ListItems.Item(i).Bold = True
      lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(11).Bold = True
      lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &HFF0000
      lv_devoluciones.ListItems.Item(i).ListSubItems(11).ForeColor = &HFF0000
   Next i
   lv_devoluciones.Refresh
   If Me.lv_devoluciones.ListItems.Count > 0 Then
      Me.lv_devoluciones.SetFocus
   End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   parametros(0) = "admcdindustrial"
   parametros(1) = "SIDAlmacenbkp"
   If cnn_devolucion_anes.State = 1 Then
      cnn_devolucion_anes.Close
   End If
   cnn_devolucion_anes.Open "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=devolucion_anes;Data Source=SQLQUEZADA2"
   
   rs.Open "select fecha, codigo, folio_anc, defecto, vcha_usuario, nombre_usuario, descripcion, cantidad, causa_original, agente, causa_real, nombre_cliente from xxvia_tb_dev_clientes_anc", cnnoracle_4, adOpenDynamic, adLockOptimistic
   
   While Not rs.EOF
         Set list_item = Me.lv_devoluciones.ListItems.Add(, , rs!codigo)
         list_item.SubItems(1) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
         list_item.SubItems(2) = 1
         list_item.SubItems(3) = IIf(IsNull(rs!causa_original), "", rs!causa_original)
         list_item.SubItems(4) = IIf(IsNull(rs!Agente), "", rs!Agente)
         list_item.SubItems(5) = IIf(IsNull(rs!nombre_cliente), "", rs!nombre_cliente)
         list_item.SubItems(6) = rs!causa_real
         list_item.SubItems(7) = rs!folio_anc
         list_item.SubItems(8) = IIf(IsNull(rs!nombre_usuario), "", rs!nombre_usuario)
         list_item.SubItems(9) = IIf(IsNull(rs!Fecha), "", rs!Fecha)
         rs.MoveNext
   Wend
   rs.Close
   
   
   Me.frm_lista.Visible = False
   Me.frm_mensaje.Visible = False
End Sub

Private Sub lv_devoluciones_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_devoluciones, ColumnHeader)
End Sub

Private Sub lv_devoluciones_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = 115 Then
      rs.Open "select * from xxvia_tb_dev_clientes_DESGLOCE where NUMERO = " + CStr(var_numero_folio_devoluciones) + " and organizacion = " + var_unidad_organizacional + " and movimiento = '" + var_clave_movimiento + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      VAR_ESTATUS = IIf(IsNull(rs!estatus), "", rs!estatus)
      rs.Close
      If VAR_ESTATUS = "I" Then
         Me.lv_lista.ListItems.Clear
         var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and enabled_flag = 'Y' ORDER BY 1"
         rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = Me.lv_lista.ListItems.Add(, , rs(0).Value)
               list_item.SubItems(1) = rs(1).Value
               rs.MoveNext
         Wend
         rs.Close
         If lv_lista.ListItems.Count > 0 Then
            Me.frm_lista.Visible = True
            Me.lv_lista.SetFocus
         Else
            MsgBox "No existen causas de devolución", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El movimiento ya no puede ser modificado", vbOKOnly, "ATENCION"
      End If
      
   End If
End Sub

Private Sub lv_devoluciones_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_devoluciones.selectedItem.Index
      If lv_devoluciones.selectedItem.SubItems(4) = "*" Then
         lv_devoluciones.selectedItem.SubItems(4) = ""
         lv_devoluciones.ListItems.Item(i).Bold = False
         lv_devoluciones.ListItems.Item(i).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(11).Bold = False
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(9).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(10).ForeColor = &H80000012
         lv_devoluciones.ListItems.Item(i).ListSubItems(11).ForeColor = &H80000012
         lv_devoluciones.Refresh
      Else
         lv_devoluciones.selectedItem.SubItems(4) = "*"
         lv_devoluciones.ListItems.Item(i).Bold = True
         lv_devoluciones.ListItems.Item(i).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).Bold = True
         lv_devoluciones.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(3).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(4).ForeColor = &HFF0000
         lv_devoluciones.ListItems.Item(i).ListSubItems(5).ForeColor = &HFF0000
         lv_devoluciones.Refresh
      End If
   End If

End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
   If KeyAscii = 13 Then
      Me.txt_defecto = Me.lv_lista.selectedItem
      Me.txt_defecto.SetFocus
      x = 0
      If x = 1 Then
      var_si = MsgBox("Se asignara la causa de devolución a los artículos seleccionados", vbYesNo, "ATENCION")
      If var_si = 6 Then
         strconsulta = "update xxvia_tb_dev_clientes_desgloce set causa_real = ?, causa_real_descripcion = ?, folio_anc = ?, fecha_asignacion = sysdate, audito = ? where rowid = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lv_lista.selectedItem)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, Me.lv_lista.selectedItem.SubItems(1))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, Me.txt_folio_ANC)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, Me.txt_usuario)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 500, var_indice)
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         
         strconsulta = "select * from xxvia_Tb_Devoluciones_clientes where movimiento = ? and numero = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_clave_movimiento)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 500, var_numero)
              .Parameters.Append parametro
         End With
         Set rsaux9 = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         var_agente = rsaux9!nombre_agente
         var_cliente = rsaux9!nombre_cliente
         
         Set list_item = Me.lv_devoluciones.ListItems.Add(, , Me.txt_codigo)
         list_item.SubItems(1) = var_descripcion
         list_item.SubItems(2) = var_cantidad
         list_item.SubItems(3) = var_causa_original
         list_item.SubItems(4) = var_agente
         list_item.SubItems(5) = var_cliente
         list_item.SubItems(6) = Me.lv_lista.selectedItem.SubItems(1)
         list_item.SubItems(7) = Me.txt_folio_ANC
         list_item.SubItems(9) = Date
         

         
         
         
         Me.txt_codigo = ""
         Me.txt_folio_ANC = ""
         Me.txt_codigo.SetFocus
         
         
         
         
         
         
         Me.frm_lista.Visible = False
      
      Else
         Me.lv_devoluciones.SetFocus
      End If
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub Text1_Change()

End Sub


Private Sub txt_codigo_GotFocus()
   Me.txt_codigo = ""
   var_encontro = 0
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Me.txt_codigo <> "" Then
          strconsulta = "SELECT A.INVENTORY_ITEM_ID, B.DESCRIPTION, cross_reference, b.segment1, nvl(a.description,'') as localizador, b.UNIT_WEIGHT, nvl(a.attribute1,1) as cantidad FROM mtl_cross_references_b A, xxvia_system_items_b B WHERE A.inventory_item_id = B.inventory_item_id AND B.organization_id = ? AND (CROSS_REFERENCE = ? OR b.segment1 = ?)"
          With comandoORA
               .ActiveConnection = cnnoracle_4
               .CommandType = adCmdText
               .CommandText = strconsulta
               Set parametro = .CreateParameter(, adNumeric, adParamInput, 3, var_unidad_organizacional)
               .Parameters.Append parametro
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
               Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
           End With
           Set rsaux8 = comandoORA.execute
           Set comandoORA = Nothing
           Set parametro = Nothing
           If Not rsaux8.EOF Then
              Me.txt_codigo.Text = IIf(IsNull(rsaux8!SEGMENT1), "", rsaux8!SEGMENT1)
              'Me.lbl_articulo = IIf(IsNull(rsaux8!SEGMENT1), "", rsaux8!SEGMENT1) + " " + IIf(IsNull(rsaux8!Description), "", rsaux8!Description)
              Me.lbl_articulo = IIf(IsNull(rsaux8!Description), "", rsaux8!Description)
           Else
              Me.txt_codigo = ""
           End If
           rsaux8.Close
       End If
       If Me.txt_codigo <> "" Then
          Me.txt_folio_ANC.SetFocus
       Else
          'MsgBox "El código leido no se encuentra en la relación", vbOKOnly, "ATENCION"
          frmmensaje.lbl_mensaje = "El artículo no existe"
          frmmensaje.Show 1
       
       End If
    End If
End Sub

Private Sub txt_defecto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.lv_lista.ListItems.Clear
      rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and enabled_flag = 'Y' ORDER BY 1"
      rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rsaux1.EOF
            Set list_item = Me.lv_lista.ListItems.Add(, , rsaux1(0).Value)
            list_item.SubItems(1) = rsaux1(1).Value
            rsaux1.MoveNext
      Wend
      rsaux1.Close
      Me.frm_lista.Visible = True
      Me.lv_lista.SetFocus
   End If
End Sub

Private Sub txt_defecto_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_defecto) Then
         Me.txt_defecto = CStr(CDbl(Me.txt_defecto))
         rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and enabled_flag = 'Y' and lookup_code = '" + CStr(CDbl(Me.txt_defecto)) + "'"
         rsaux1.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.lbl_defecto = IIf(IsNull(rsaux1!nombre), "", rsaux1!nombre)
            If Me.txt_defecto <> "" Then
               If Me.txt_folio_ANC <> "" Then
                  If Me.txt_codigo <> "" Then
                     If Me.txt_usuario <> "" Then
                        strconsulta = "insert into xxvia_tb_dev_clientes_anc (fecha, codigo, folio_anc, defecto, vcha_usuario, nombre_usuario, descripcion, cantidad, causa_original, agente, causa_real, nombre_cliente, folio_devolucion) values (sysdate,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_folio_ANC)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_defecto)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_usuario)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lbl_usuario)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lbl_articulo)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, 1)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lbl_defecto_original)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lbl_nombre_agente)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lbl_defecto)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.lbl_nombre_cliente)
                             .Parameters.Append parametro
                             Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, var_folio_Devolucion)
                             .Parameters.Append parametro
                        End With
                        Set rsaux8 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        Set list_item = Me.lv_devoluciones.ListItems.Add(, , Me.txt_codigo)
                        list_item.SubItems(1) = Me.lbl_articulo
                        list_item.SubItems(2) = 1
                        list_item.SubItems(3) = Me.lbl_defecto_original
                        list_item.SubItems(4) = Me.lbl_nombre_agente
                        list_item.SubItems(5) = Me.lbl_nombre_cliente
                        list_item.SubItems(6) = Me.lbl_defecto
                        list_item.SubItems(7) = Me.txt_folio_ANC
                        list_item.SubItems(8) = Me.lbl_usuario
                        Me.txt_usuario = ""
                        Me.txt_defecto = ""
                        Me.txt_folio_ANC = ""
                        Me.txt_codigo = ""
                        Me.txt_usuario.SetFocus
                     Else
                        frmmensaje.lbl_mensaje = "El usuario no existe."
                        frmmensaje.Show 1
                     End If
                  Else
                     frmmensaje.lbl_mensaje = "Código de artículo incorrecto."
                     frmmensaje.Show 1
                  End If
              Else
                 frmmensaje.lbl_mensaje = "Folio de ANC incorrecto."
                 frmmensaje.Show 1
              End If
           Else
              frmmensaje.lbl_mensaje = "Usuario incorrecto."
              frmmensaje.Show 1
           End If
         Else
            Me.txt_defecto = ""
            Me.lbl_defecto = ""
         End If
         rsaux1.Close
      Else
         Me.txt_defecto = ""
         frmmensaje.lbl_mensaje = "Defecto incorrecto"
         frmmensaje.Show 1
      End If
      

   End If
End Sub

Private Sub txt_filtro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         If Me.txt_filtro <> "" Then
            Me.lv_lista.ListItems.Clear
            var_cadena = "select lookup_code as CODIGO, meaning as NOMBRE, description as DESCRIPCION From FND_LOOKUP_VALUES_VL where lookup_type = 'CREDIT_MEMO_REASON' and meaning like '%?%' and enabled_flag = 'Y' ORDER BY 1"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_filtro)
                 .Parameters.Append parametro
            End With
            Set rs = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            While Not rs.EOF
                  Set list_item = lv_lista.ListItems.Add(, , rs!codigo)
                  list_item.SubItems(1) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
                  rs.MoveNext
            Wend
            rs.Close
         End If
      End If
End Sub

Private Sub txt_folio_ANC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_codigo <> "" Then
         If IsNumeric(Me.txt_folio_ANC) Then
            If CDbl(Me.txt_folio_ANC) > 0 Then
               rs.Open "SELECT * FROM TB_DEVOLUCIONES WHERE LIMITE_INFERIOR <= " + Me.txt_folio_ANC + " AND LIMITE_superior >= " + Me.txt_folio_ANC + " AND CODIGO = '" + Me.txt_codigo + "'", cnn_devolucion_anes, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                  Me.lbl_anc = "AGENTE: " + IIf(IsNull(rs!nombre_agente), "", rs!nombre_agente)
                  Me.lbl_defecto_original = IIf(IsNull(rs!tipo_Devolucion_2), "", rs!tipo_Devolucion_2)
                  Me.lbl_nombre_agente = IIf(IsNull(rs!nombre_agente), "", rs!nombre_agente)
                  lbl_nombre_cliente = IIf(IsNull(rs!nombre_cliente), "", rs!nombre_cliente)
                  var_folio_Devolucion = CStr(IIf(IsNull(rs!NUMERO), "", rs!NUMERO))
                  Me.txt_defecto.SetFocus
               Else
                  Me.txt_folio_ANC = ""
                  'MsgBox "No existe la relación del folio ANC con el código del artículo", vbOKOnly, "ATENCION"
                  frmmensaje.lbl_mensaje = "No existe la relación del folio ANC con el código del artículo"
                  frmmensaje.Show 1
                  Me.txt_defecto = ""
                  Me.lbl_defecto_original = ""
                  Me.lbl_nombre_agente = ""
                  lbl_nombre_cliente = ""
                  var_folio_Devolucion = ""
               End If
               rs.Close
            Else
               'MsgBox "Código de ANC incorrecto", vbOKOnly, "ATENCION"

               frmmensaje.lbl_mensaje = "Código de ANC incorrecto"
               frmmensaje.Show 1
            
            End If
         Else
            'MsgBox "Número de folo de ANC incorrecto.", vbOKOnly, "ATENCION"
            frmmensaje.lbl_mensaje = "Código de ANC incorrecto"
            frmmensaje.Show 1
            
         End If
      Else
         'MsgBox "Código de artículo incorrecto.", vbOKOnly, "ATENCION"
         frmmensaje.lbl_mensaje = "Código de artículo incorrecto."
         frmmensaje.Show 1
      End If
      If Me.txt_folio_ANC <> "" Then
         Me.txt_defecto.SetFocus
      End If
   End If
End Sub

Private Sub txt_usuario_Change()
   Me.lbl_usuario = ""
   Me.lbl_usuario = ""
   Me.lbl_articulo = ""
   Me.lbl_defecto = ""
   Me.lbl_anc = ""
   Me.lbl_defecto_original = ""
   Me.lbl_nombre_agente = ""
   lbl_nombre_cliente = ""
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.lbl_usuario = ""
      Me.lbl_articulo = ""
      Me.lbl_defecto = ""
      Me.lbl_anc = ""
      Me.lbl_defecto_original = ""
      Me.lbl_nombre_agente = ""
      lbl_nombre_cliente = ""
      If Me.txt_usuario <> "" Then
         rs.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + Me.txt_usuario + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_codigo.SetFocus
            Me.lbl_usuario = IIf(IsNull(rs!vcha_usu_nombre), "", rs!vcha_usu_nombre) + " " + IIf(IsNull(rs!vcha_usu_apellidos), "", rs!vcha_usu_apellidos)
            
         Else
            frmmensaje.lbl_mensaje = "El usuario no existe"
            frmmensaje.Show 1

         End If
         rs.Close
      End If
   End If
End Sub
