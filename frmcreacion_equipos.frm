VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcreacion_equipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creaci?n de Equipos de Embarques"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   5355
      TabIndex        =   14
      Top             =   420
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   15
         Top             =   480
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3228
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
         TabIndex        =   16
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmcreacion_equipos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11145
      Picture         =   "frmcreacion_equipos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   75
      TabIndex        =   6
      Top             =   270
      Width           =   11445
   End
   Begin VB.Frame Frame3 
      ClipControls    =   0   'False
      Height          =   6765
      Left            =   105
      TabIndex        =   4
      Top             =   420
      Width           =   4500
      Begin VB.OptionButton equipo_13 
         Caption         =   "Equipo 13"
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
         Left            =   210
         TabIndex        =   31
         Top             =   5724
         Width           =   3975
      End
      Begin VB.OptionButton equipo_14 
         Caption         =   "Equipo 14"
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
         Left            =   210
         TabIndex        =   30
         Top             =   6165
         Width           =   3975
      End
      Begin VB.OptionButton equipo_7 
         Caption         =   "Equipo 7"
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
         Left            =   210
         TabIndex        =   29
         Top             =   3102
         Width           =   3975
      End
      Begin VB.OptionButton equipo_8 
         Caption         =   "Equipo 8"
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
         Left            =   210
         TabIndex        =   28
         Top             =   3539
         Width           =   3975
      End
      Begin VB.OptionButton equipo_9 
         Caption         =   "Equipo 9"
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
         Left            =   210
         TabIndex        =   27
         Top             =   3976
         Width           =   3975
      End
      Begin VB.OptionButton equipo_10 
         Caption         =   "Equipo 10"
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
         Left            =   210
         TabIndex        =   26
         Top             =   4413
         Width           =   3975
      End
      Begin VB.OptionButton equipo_11 
         Caption         =   "Equipo 11"
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
         Left            =   210
         TabIndex        =   25
         Top             =   4850
         Width           =   3975
      End
      Begin VB.OptionButton equipo_12 
         Caption         =   "Equipo 12"
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
         Left            =   210
         TabIndex        =   24
         Top             =   5287
         Width           =   3975
      End
      Begin VB.OptionButton equipo_6 
         Caption         =   "Equipo 6"
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
         Left            =   210
         TabIndex        =   23
         Top             =   2665
         Width           =   3975
      End
      Begin VB.OptionButton equipo_5 
         Caption         =   "Equipo 5"
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
         Left            =   210
         TabIndex        =   21
         Top             =   2228
         Width           =   3975
      End
      Begin VB.OptionButton equipo_4 
         Caption         =   "Equipo 4"
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
         Left            =   210
         TabIndex        =   20
         Top             =   1791
         Width           =   3975
      End
      Begin VB.OptionButton equipo_3 
         Caption         =   "Equipo 3"
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
         Left            =   210
         TabIndex        =   19
         Top             =   1354
         Width           =   3975
      End
      Begin VB.OptionButton equipo_2 
         Caption         =   "Equipo 2"
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
         Left            =   210
         TabIndex        =   18
         Top             =   917
         Width           =   3975
      End
      Begin VB.OptionButton equipo_1 
         Caption         =   "Equipo 1"
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
         Left            =   225
         TabIndex        =   17
         Top             =   465
         Width           =   3975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   " Equipos "
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   30
         TabIndex        =   5
         Top             =   120
         Width           =   4425
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   4695
      TabIndex        =   2
      Top             =   3795
      Width           =   6810
      Begin VB.TextBox txt_orden_surtido 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   450
         Width           =   1200
      End
      Begin MSComctlLib.ListView lv_ordenes_surtido 
         Height          =   2145
         Left            =   120
         TabIndex        =   13
         Top             =   810
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   3784
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
            Text            =   "N?mero"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agente"
            Object.Width           =   6985
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lbl_cantidad 
         Alignment       =   1  'Right Justify
         Caption         =   "999999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4425
         TabIndex        =   22
         Top             =   2985
         Width           =   2250
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Ordenes de surtido "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   4695
      TabIndex        =   0
      Top             =   420
      Width           =   6810
      Begin VB.TextBox txt_nombre_persona 
         Height          =   315
         Left            =   1365
         TabIndex        =   10
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txt_persona 
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   480
         Width           =   1185
      End
      Begin MSComctlLib.ListView lv_personas 
         Height          =   2430
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   4286
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   9631
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   " Colaboradores del equipo"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmcreacion_equipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_fecha_numero As Double
Dim var_equipo As Integer

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub equipo_1_Click()
   var_equipo = 1
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from ar_collectors where collector_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!Name), "", rsaux!Name)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
   
End Sub

Private Sub equipo_10_Click()
   var_equipo = 10
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_11_Click()
   var_equipo = 11
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_12_Click()
   var_equipo = 12
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_13_Click()
   var_equipo = 13
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_14_Click()
   var_equipo = 14
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_2_Click()
   var_equipo = 2
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_3_Click()
   var_equipo = 3
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_4_Click()
   var_equipo = 4
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_5_Click()
   var_equipo = 5
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_6_Click()
   var_equipo = 6
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_7_Click()
   var_equipo = 7
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_8_Click()
   var_equipo = 8
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub equipo_9_Click()
   var_equipo = 9
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from tb_agentes where vcha_age_agente_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   frm_lista.Visible = False
   equipo_1.Value = True
   var_mes = CStr(Month(Date))
   var_dia = CStr(Day(Date))
   If Len(var_mes) = 1 Then
      var_mes = "0" + var_mes
   End If
   If Len(var_dia) = 1 Then
      var_dia = "0" + var_dia
   End If
   var_fecha_numero = CDbl(CStr(Year(Date)) + var_mes + var_dia)
   var_equipo = 1
   lv_personas.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = lv_personas.ListItems.Add(, , rs!vcha_per_personal_id)
            rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            rs.MoveNext
      Wend
   End If
   rs.Close
   lbl_cantidad = "0"
   Me.lv_ordenes_surtido.ListItems.Clear
   rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      While Not rs.EOF
            Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , rs!inte_ors_orden_surtido)
            rsaux.Open "select * from ar_collectors where collector_id = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               list_item.SubItems(1) = IIf(IsNull(rsaux!Name), "", rsaux!Name)
            Else
               list_item.SubItems(1) = ""
            End If
            rsaux.Close
            list_item.SubItems(2) = Format(IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad), "###,###,##0.00")
            lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0.00")
            rs.MoveNext
      Wend
   End If
   rs.Close
   If Me.lv_ordenes_surtido.ListItems.Count > 11 Then
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3710
   Else
      Me.lv_ordenes_surtido.ColumnHeaders(2).Width = 3960
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If lv_lista.ListItems.Count > 0 Then
      Me.txt_persona = lv_lista.selectedItem
      Me.txt_nombre_persona = lv_lista.selectedItem.SubItems(1)
      Me.txt_persona.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub lv_ordenes_surtido_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      var_mes = CStr(Month(Date))
      var_dia = CStr(Day(Date))
      If Len(var_mes) = 1 Then
         var_mes = "0" + var_mes
      End If
      If Len(var_dia) = 1 Then
         var_dia = "0" + var_dia
      End If
      var_fecha_numero = CDbl(CStr(Year(Date)) + var_mes + var_dia)
      If Me.equipo_1.Value = True Then
         var_equipo = 1
      End If
      If Me.equipo_2.Value = True Then
         var_equipo = 2
      End If
      If Me.equipo_3.Value = True Then
         var_equipo = 3
      End If
      If Me.equipo_4.Value = True Then
         var_equipo = 4
      End If
      If Me.equipo_5.Value = True Then
         var_equipo = 5
      End If
      var_si = MsgBox("?Desea eliminar el registro?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "DELETE from TB_DETALLE_EQUIPOS_ORDEN_SURTIDO where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_ors_orden_surtido = " + Me.lv_ordenes_surtido.selectedItem, cnn, adOpenDynamic, adLockOptimistic
         Me.lbl_cantidad = Format(CDbl(lbl_cantidad) - CDbl(Me.lv_ordenes_surtido.selectedItem.SubItems(2)), "###,###,##0.00")
         Me.lv_ordenes_surtido.ListItems.Remove (lv_ordenes_surtido.selectedItem.Index)
      End If
   End If
End Sub

Private Sub lv_personas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      var_mes = CStr(Month(Date))
      var_dia = CStr(Day(Date))
      If Len(var_mes) = 1 Then
         var_mes = "0" + var_mes
      End If
      If Len(var_dia) = 1 Then
         var_dia = "0" + var_dia
      End If
      var_fecha_numero = CDbl(CStr(Year(Date)) + var_mes + var_dia)
      If Me.equipo_1.Value = True Then
         var_equipo = 1
      End If
      If Me.equipo_2.Value = True Then
         var_equipo = 2
      End If
      If Me.equipo_3.Value = True Then
         var_equipo = 3
      End If
      If Me.equipo_4.Value = True Then
         var_equipo = 4
      End If
      If Me.equipo_5.Value = True Then
         var_equipo = 5
      End If
      var_si = MsgBox("?Desea eliminar el registro?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "DELETE from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo) + " AND vcha_per_personal_id = '" + Me.lv_personas.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
         lv_personas.ListItems.Remove (lv_personas.selectedItem.Index)
      End If
   End If
End Sub

Private Sub txt_nombre_persona_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      If Trim(Me.txt_persona) <> "" Then
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and vcha_per_personal_id = '" + Me.txt_persona + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            MsgBox "El colaborador ya se encuentra en el equipo " + CStr(rs!inte_equ_equipo), vbOKOnly, "ATENCION"
         Else
            If equipo_1.Value = True Then
               var_equipo = 1
            End If
            If equipo_2.Value = True Then
               var_equipo = 2
            End If
            If equipo_3.Value = True Then
               var_equipo = 3
            End If
            If equipo_4.Value = True Then
               var_equipo = 4
            End If
            If equipo_5.Value = True Then
               var_equipo = 5
            End If
            If equipo_6.Value = True Then
               var_equipo = 6
            End If
            If equipo_7.Value = True Then
               var_equipo = 7
            End If
            If equipo_8.Value = True Then
               var_equipo = 8
            End If
            If equipo_9.Value = True Then
               var_equipo = 9
            End If
            If equipo_10.Value = True Then
               var_equipo = 10
            End If
            If equipo_11.Value = True Then
               var_equipo = 11
            End If
            If equipo_12.Value = True Then
               var_equipo = 12
            End If
            If equipo_13.Value = True Then
               var_equipo = 13
            End If
            If equipo_14.Value = True Then
               var_equipo = 14
            End If
            rsaux.Open "insert into tb_detalle_equipos_personal (inte_equ_numero, inte_equ_equipo, vcha_per_personal_id) values (" + CStr(var_fecha_numero) + ", " + CStr(var_equipo) + "," + Me.txt_persona + ")", cnn, adOpenDynamic, adLockOptimistic
            Set list_item = lv_personas.ListItems.Add(, , Me.txt_persona)
            list_item.SubItems(1) = Me.txt_nombre_persona
         End If
         rs.Close
         Me.txt_persona.SetFocus
         Me.txt_persona = ""
         Me.txt_nombre_persona = ""
      End If
   End If
End Sub

Private Sub txt_orden_surtido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_orden_surtido) Then
         rsaux4.Open "select * from tb_Detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_ors_orden_surtido = " + Me.txt_orden_surtido, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux4.EOF Then
            MsgBox "La orden de surtido se encuentra en el equipo " + CStr(rsaux4!inte_equ_equipo), vbOKOnly, "ATENCION"
         Else
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_cadena = "SELECT E.collector_id AS VCHA_AGE_AGENTE_ID, E.name AS VCHA_AGE_NOMBRE, A.released_status, sum(requested_quantity) AS FLOA_ORS_CANTIDAD_SURTIR from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER = '" + Me.txt_orden_surtido + "'"
            var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id  AND RELEASED_STATUS = 'Y' GROUP BY E.collector_id, E.name, A.released_status"
            rs.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_agente = rs!VCHA_AGE_AGENTE_ID
               If equipo_1.Value = True Then
                  var_equipo = 1
               End If
               If equipo_2.Value = True Then
                  var_equipo = 2
               End If
               If equipo_3.Value = True Then
                  var_equipo = 3
               End If
               If equipo_4.Value = True Then
                  var_equipo = 4
               End If
               If equipo_5.Value = True Then
                  var_equipo = 5
               End If
               If equipo_6.Value = True Then
                  var_equipo = 6
               End If
               If equipo_7.Value = True Then
                  var_equipo = 7
               End If
               If equipo_8.Value = True Then
                  var_equipo = 8
               End If
               If equipo_9.Value = True Then
                  var_equipo = 9
               End If
               If equipo_10.Value = True Then
                  var_equipo = 10
               End If
               If equipo_11.Value = True Then
                  var_equipo = 11
               End If
               If equipo_12.Value = True Then
                  var_equipo = 12
               End If
               If equipo_13.Value = True Then
                  var_equipo = 13
               End If
               If equipo_14.Value = True Then
                  var_equipo = 14
               End If
               
               var_cantidad = rs!FLOA_ORS_CANTIDAD_SURTIR
               If var_cantidad < 0 Then
                  var_cantidad = 0
               End If
               Set list_item = Me.lv_ordenes_surtido.ListItems.Add(, , Me.txt_orden_surtido)
               list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
               list_item.SubItems(2) = Format(var_cantidad, "###,###,##0.00")
               lbl_cantidad = Format(CStr(CDbl(lbl_cantidad) + var_cantidad), "###,###,##0.00")
               var_cantidad_surtida = 0
               var_cantidad_empacada = 0
               var_cantidad_negada = 0
               rsaux2.Open "insert into TB_DETALLE_EQUIPOS_ORDEN_SURTIDO (inte_equ_numero, inte_equ_equipo, inte_ors_orden_surtido, vcha_age_Agente_id, floa_ors_cantidad, floa_ors_cantidad_surtida, floa_ors_cantidad_negada) values (" + CStr(var_fecha_numero) + "," + CStr(var_equipo) + "," + Me.txt_orden_surtido + ",'" + CStr(var_agente) + "'," + CStr(var_cantidad) + ", " + CStr(var_cantidad_surtida + var_cantidad_empacada) + ", " + CStr(var_cantidad_negada) + ")", cnn, adOpenDynamic, adLockOptimistic
            Else
               MsgBox "La orden de surtido no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If
         rsaux4.Close
         Me.txt_orden_surtido = ""
      Else
         MsgBox "Orden de surtido incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_persona_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_personal order by vcha_per_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_per_personal_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_per_nombre), "", rs!vcha_per_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "COLABORADORES"
      VAR_TIPO_LISTA = 5
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

Private Sub txt_persona_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_persona_LostFocus()
   If Trim(Me.txt_persona) <> "" Then
      rs.Open "select * from tb_personal where vcha_per_personaL_id = '" + Me.txt_persona + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_persona = IIf(IsNull(rs!vcha_per_nombre), "", rs!vcha_per_nombre)
      Else
         MsgBox "Clave de colaborador incorrecto", vbOKOnly, "ATENCION"
         Me.txt_persona = ""
      End If
      rs.Close
   End If
End Sub
