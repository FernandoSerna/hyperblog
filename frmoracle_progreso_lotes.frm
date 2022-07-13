VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_progreso_lotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progreso de lotes"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   20280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   20280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_procesando 
      Height          =   855
      Left            =   8400
      TabIndex        =   28
      Top             =   4680
      Width           =   3855
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Procesando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3540
      Top             =   15
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   75
      TabIndex        =   3
      Top             =   135
      Width           =   20100
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   19125
         Top             =   255
      End
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   19560
         Picture         =   "frmoracle_progreso_lotes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Actualiza Grafica"
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   435
         Left            =   16095
         TabIndex        =   17
         Top             =   330
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_porcentaje 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   17490
         TabIndex        =   10
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label lbl_cantidad_surtida 
         AutoSize        =   -1  'True
         Caption         =   "12365"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   14355
         TabIndex        =   9
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad surtida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11265
         TabIndex        =   8
         Top             =   270
         Width           =   3210
      End
      Begin VB.Label lbl_surtir 
         AutoSize        =   -1  'True
         Caption         =   "12312"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   8310
         TabIndex        =   7
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad por surtir:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4890
         TabIndex        =   6
         Top             =   270
         Width           =   3360
      End
      Begin VB.Label lbl_embarque 
         AutoSize        =   -1  'True
         Caption         =   "123456"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2115
         TabIndex        =   5
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   165
         TabIndex        =   4
         Top             =   270
         Width           =   1920
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9630
      Left            =   75
      TabIndex        =   1
      Top             =   945
      Width           =   20100
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   15
         TabIndex        =   27
         Top             =   8310
         Width           =   21600
      End
      Begin MSComctlLib.ListView lv_pantalla 
         CausesValidation=   0   'False
         Height          =   7230
         Left            =   60
         TabIndex        =   18
         Top             =   150
         Visible         =   0   'False
         Width           =   19920
         _ExtentX        =   35137
         _ExtentY        =   12753
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lt."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pr."
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pedidas"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Leidas"
            Object.Width           =   2152
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Máquina"
            Object.Width           =   3440
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Usuario"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "   %"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Tiempo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Estatus"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Aduana"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView lv_pedidos 
         CausesValidation=   0   'False
         Height          =   7260
         Left            =   60
         TabIndex        =   2
         Top             =   150
         Width           =   19965
         _ExtentX        =   35216
         _ExtentY        =   12806
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pedido"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lt."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pr."
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cliente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pedidas"
            Object.Width           =   2152
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Leidas"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Máquina"
            Object.Width           =   3440
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Usuario"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "   %"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Tiempo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Estatus"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Aduana"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Consecutivo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Aduana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   26
         Top             =   8475
         Width           =   1320
      End
      Begin VB.Label lbl_tiempo_transcurrido_aduana 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo transcurrido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   14820
         TabIndex        =   25
         Top             =   8985
         Width           =   3660
      End
      Begin VB.Label lbl_tiempo_aduana 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   18495
         TabIndex        =   24
         Top             =   8985
         Width           =   1485
      End
      Begin VB.Label lbl_hora_inicio_aduana 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   2160
         TabIndex        =   23
         Top             =   9015
         Width           =   1485
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hora inicio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   22
         Top             =   8970
         Width           =   2010
      End
      Begin VB.Label lbl_hora_final_aduana 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   10110
         TabIndex        =   21
         Top             =   9000
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Hora final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8250
         TabIndex        =   20
         Top             =   8970
         Width           =   1800
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Embarques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   90
         TabIndex        =   19
         Top             =   7380
         Width           =   2010
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hora final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8235
         TabIndex        =   16
         Top             =   7875
         Width           =   1800
      End
      Begin VB.Label lbl_hora_final 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   10095
         TabIndex        =   15
         Top             =   7905
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hora inicio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   90
         TabIndex        =   14
         Top             =   7875
         Width           =   2010
      End
      Begin VB.Label lbl_hora_inicio 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   2145
         TabIndex        =   13
         Top             =   7920
         Width           =   1485
      End
      Begin VB.Label lbl_tiempo 
         AutoSize        =   -1  'True
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   18480
         TabIndex        =   12
         Top             =   7890
         Width           =   1485
      End
      Begin VB.Label lbl_tiempo_transcurrido 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo transcurrido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   14805
         TabIndex        =   11
         Top             =   7890
         Width           =   3660
      End
   End
End
Attribute VB_Name = "frmoracle_progreso_lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim var_veces As Integer

Private Sub Command1_Click()
   Me.Timer1.Enabled = False
   'Me.lv_pantalla.Visible = True
   var_fecha_inicio_pantalla = Now
   Me.lbl_tiempo = ""
   Me.lbl_tiempo_transcurrido = ""
   Me.lbl_tiempo_aduana = ""
   Me.lbl_tiempo_transcurrido_aduana = ""
   Me.lbl_surtir = "0.00"
   Me.lbl_cantidad_surtida = "0.00"
   Me.lbl_porcentaje = "0.00%"
   Me.lbl_embarque = Format(var_embarque_lotes, "#########")
   Me.frm_procesando.Visible = True
   
   
   'rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(var_embarque_lotes) + " order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
   rs.Open "select pedido from tb_oracle_pedidos_asignados_embarques where embarque = " + CStr(var_embarque_lotes) + " order by orden_pedido, pedido", cnn, adOpenDynamic, adLockOptimistic
   var_Cadena_pedidos = ""
   While Not rs.EOF
         If var_Cadena_pedidos = "" Then
            var_Cadena_pedidos = CStr(rs!pedido)
         Else
            var_Cadena_pedidos = var_Cadena_pedidos + "," + CStr(rs!pedido)
         End If
         rs.MoveNext
   Wend
   rs.Close
   lv_pedidos.ListItems.Clear
   'While Not rs.EOF
         strconsulta = "select nvl(estatus_lote,0) as estatus_lote, source_header_number, lote, sum(src_requested_quantity) as cantidad, max(maquina) as maquina, max(usuario) as usuario, (select sum(floa_sal_Cantidad_leida) from xxvia_Tb_salidas_cajas where source_header_number = a.source_header_number and lote = a.lote) as cantidad_leida from xxvia_tb_pedidos_divididos a where source_header_number = ? group by nvl(estatus_lote,0), source_header_number, lote order by lote"
         strconsulta = "select nvl(estatus_lote,0) as estatus_lote, source_header_number, lote, sum(src_requested_quantity) as cantidad, max(maquina) as maquina, max(usuario) as usuario, 0 as cantidad_leida from xxvia_tb_pedidos_divididos a where source_header_number = ? group by nvl(estatus_lote,0), source_header_number, lote order by lote"
         
         strconsulta = "select nvl(estatus_lote,0) as estatus_lote, source_header_number, lote, sum(src_requested_quantity) as cantidad, max(maquina) as maquina, max(usuario) as usuario, 0 as cantidad_leida from xxvia_tb_pedidos_divididos a where source_header_number in (?) group by nvl(estatus_lote,0), source_header_number, lote order by lote"
         'With comandoORA
         '     .ActiveConnection = cnnoracle_4
         '     .CommandType = adCmdText
         '     .CommandText = strconsulta
         '     'Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!pedido)
         '     Set parametro = .CreateParameter(, adVarChar, adParamInput, 1000, var_cadena_pedidos)
         '     .Parameters.Append parametro
         'End With
         'Set rsaux6 = comandoORA.execute
         'Set comandoORA = Nothing
         'Set parametro = Nothing
         
         strconsulta = "select nvl(estatus_lote,0) as estatus_lote, source_header_number, lote, sum(src_requested_quantity) as cantidad, max(maquina) as maquina, max(usuario) as usuario, 0 as cantidad_leida from xxvia_tb_pedidos_divididos a where source_header_number in (" + var_Cadena_pedidos + ") group by nvl(estatus_lote,0), source_header_number, lote order by lote"
         rsaux6.Open strconsulta, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux6.EOF
               Me.Refresh
               rs.Open "select DISTINCT PEDIDO, CLIENTE, PIEZAS, agente, nombre_agente, orden_pedido, volumen from tb_oracle_pedidos_asignados_embarques where pedido = " + CStr(rsaux6!source_header_number) + " order by orden_pedido", cnn, adOpenDynamic, adLockOptimistic
               Set list_item = lv_pedidos.ListItems.Add(, , rs!pedido)
               list_item.SubItems(1) = rsaux6!lote
               'list_item.SubItems(2) = rs!orden_pedido
               list_item.SubItems(2) = Format(rs!orden_pedido, "@@@")
               If rs!Cliente = "VIANNEY TEXTIL HOGAR SA DE CV" Then
                  var_cliente = rs!nombre_agente
               Else
                  var_cliente = rs!Cliente
               End If
               list_item.SubItems(3) = var_cliente
               list_item.SubItems(4) = rsaux6!cantidad
               x = 1
               If x = 0 Then
                  strconsulta = "select source_header_number, lote, sum(floa_sal_Cantidad_leida) as cantidad from xxvia_tb_Salidas_cajas where source_header_number = ? and lote  = ? group by source_header_number, lote order by lote"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!pedido)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux6!lote)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux7.EOF Then
                     list_item.SubItems(5) = IIf(IsNull(rsaux7!cantidad), 0, rsaux7!cantidad)
                  Else
                     list_item.SubItems(5) = 0
                  End If
                  rsaux7.Close
               Else
                   rsaux7.Open "select cantidad from tb_oracle_suma_lotes where pedido = " + CStr(rs!pedido) + " and lote  = " + CStr(rsaux6!lote), cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux7.EOF Then
                      list_item.SubItems(5) = IIf(IsNull(rsaux7!cantidad), 0, rsaux7!cantidad)
                   Else
                      list_item.SubItems(5) = 0
                   End If
                   rsaux7.Close
               End If
               
               list_item.SubItems(6) = IIf(IsNull(rsaux6!maquina), "    ", rsaux6!maquina)
               rsaux7.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + IIf(IsNull(rsaux6!USUARIO), "", rsaux6!USUARIO) + "'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
                  list_item.SubItems(7) = IIf(IsNull(rsaux7!vcha_usu_nombre), 0, rsaux7!vcha_usu_nombre) + " " + IIf(IsNull(rsaux7!vcha_usu_apellidos), 0, rsaux7!vcha_usu_apellidos)
               Else
                  list_item.SubItems(7) = "    "
               End If
               rsaux7.Close
               If rsaux6!cantidad = IIf(IsNull(rsaux6!cantidad_leida), 0, rsaux6!cantidad_leida) Then
                   var_porcentaje = 100
               Else
                   var_porcentaje = (CDbl(list_item.SubItems(5)) * 100) / CDbl(list_item.SubItems(4))
               End If
               list_item.SubItems(8) = Format(var_porcentaje, "##0.00")
               list_item.SubItems(9) = ""
               If rsaux6!estatus_lote = 0 Then
                  var_estatus_lote = ""
               Else
                  var_estatus_lote = "Cerrado"
               End If
               list_item.SubItems(10) = var_estatus_lote
               list_item.SubItems(11) = ""
               Me.Refresh
               rs.Close
               rsaux6.MoveNext
        
         Wend
         rsaux6.Close
         'rs.MoveNext
   
   
   'Wend
   'rs.Close
   var_cerrado = 1
            For var_i = 1 To lv_pedidos.ListItems.Count
                lv_pedidos.ListItems(var_i).Selected = True
                
                
                
                If Trim(Me.lv_pedidos.selectedItem.SubItems(6)) = "" Then
                   rs.Open "select * from TB_ORACLE_TIEMPO_POR_LOTE where pedido = " + Me.lv_pedidos.selectedItem + " and lote  = " + Me.lv_pedidos.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                   If Not rs.EOF Then
                      rsaux7.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + IIf(IsNull(rs!USUARIO), "", rs!USUARIO) + "'", cnn, adOpenDynamic, adLockOptimistic
                      If Not rsaux7.EOF Then
                         VAR_USUARIO = IIf(IsNull(rsaux7!vcha_usu_nombre), 0, rsaux7!vcha_usu_nombre) + " " + IIf(IsNull(rsaux7!vcha_usu_apellidos), 0, rsaux7!vcha_usu_apellidos)
                      Else
                         VAR_USUARIO = ""
                      End If
                      rsaux7.Close
                      var_maquina = IIf(IsNull(rs!maquina), "", rs!maquina)
                   Else
                      var_maquina = ""
                      VAR_USUARIO = ""
                   End If
                   rs.Close
                   Me.lv_pedidos.selectedItem.SubItems(6) = var_maquina
                   Me.lv_pedidos.selectedItem.SubItems(7) = VAR_USUARIO
                End If
                Me.lbl_surtir = Format(CDbl(Me.lbl_surtir) + CDbl(Me.lv_pedidos.selectedItem.SubItems(4)), "###,###,##0.00")
                Me.lbl_cantidad_surtida = Format(CDbl(Me.lbl_cantidad_surtida) + CDbl(Me.lv_pedidos.selectedItem.SubItems(5)), "###,###,##0.00")
                Me.lbl_porcentaje = Format((CDbl(Me.lbl_cantidad_surtida) * 100) / CDbl(Me.lbl_surtir), "##0.00") + "%"
                rsaux7.Open "select hora_inicio, hora_final, cast(ISNULL(hora_final,getdate()) - hora_inicio as time) as tiempo from TB_ORACLE_TIEMPO_POR_LOTE where pedido = " + Me.lv_pedidos.selectedItem + " and lote  = " + Me.lv_pedidos.selectedItem.SubItems(1), cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux7.EOF Then
                   Me.lv_pedidos.selectedItem.SubItems(9) = Mid(CStr(IIf(IsNull(rsaux7!tiempo), "", rsaux7!tiempo)), 1, 8)
                End If
                rsaux7.Close
                lv_pedidos.ListItems.Item(var_i).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(1).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(2).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(2).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(3).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(3).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(4).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(4).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(5).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(5).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(6).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(6).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(7).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(7).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(8).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(8).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(9).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(9).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(10).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(10).ForeColor = &HFF&
                lv_pedidos.ListItems(var_i).ListSubItems(11).Bold = True
                lv_pedidos.ListItems(var_i).ListSubItems(11).ForeColor = &HFF&
                lv_pedidos.selectedItem.Bold = True
                If Me.lv_pedidos.selectedItem.SubItems(10) <> "Cerrado" Then
                   var_cerrado = 0
                End If
                If (lv_pedidos.selectedItem.SubItems(8) * 1) > 25 Then
                   lv_pedidos.ListItems.Item(var_i).ForeColor = &HFF0000
                   'lv_pedidos.ListItems(var_i).ListSubItems(1).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(1).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(2).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(2).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(3).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(3).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(4).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(4).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(5).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(5).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(6).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(6).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(7).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(7).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(8).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(8).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(9).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(9).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(10).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(10).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(11).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(11).ForeColor = &HFF0000
                   lv_pedidos.selectedItem.Bold = True
                End If
                If (lv_pedidos.selectedItem.SubItems(8) * 1) > 50 Or Trim(lv_pedidos.selectedItem.SubItems(9)) <> "" Then
                   lv_pedidos.ListItems.Item(var_i).ForeColor = &HFF0000
                   'lv_pedidos.ListItems(var_i).ListSubItems(1).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(1).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(2).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(2).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(3).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(3).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(4).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(4).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(5).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(5).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(6).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(6).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(7).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(7).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(8).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(8).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(9).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(9).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(10).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(10).ForeColor = &HFF0000
                   lv_pedidos.ListItems(var_i).ListSubItems(11).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(11).ForeColor = &HFF0000
                   lv_pedidos.selectedItem.Bold = True
                End If
                If (lv_pedidos.selectedItem.SubItems(8) * 1) = 100 Or lv_pedidos.selectedItem.SubItems(10) = "Cerrado" Then
                   lv_pedidos.ListItems.Item(var_i).ForeColor = &HC000&
                   'lv_pedidos.ListItems(var_i).ListSubItems(1).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(1).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(2).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(2).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(3).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(3).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(4).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(4).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(5).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(5).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(6).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(6).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(7).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(7).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(8).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(8).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(9).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(9).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(10).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(10).ForeColor = &HC000&
                   lv_pedidos.ListItems(var_i).ListSubItems(11).Bold = True
                   lv_pedidos.ListItems(var_i).ListSubItems(11).ForeColor = &HC000&
                   lv_pedidos.selectedItem.Bold = True
               End If
            Next var_i
            
            var_pedidos_embarque = ""
            rsaux7.Open "select pedido from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = " + Me.lbl_embarque, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux7.EOF
                  If var_pedidos_embarque = "" Then
                     var_pedidos_embarque = CStr(rsaux7!pedido)
                  Else
                     var_pedidos_embarque = var_pedidos_embarque + "," + CStr(rsaux7!pedido)
                  End If
                  rsaux7.MoveNext
            Wend
            rsaux7.Close
            
            If var_cerrado = 0 Then
               rsaux7.Open "select  min(hora_inicio) as hora_inicio from TB_ORACLE_TIEMPO_POR_LOTE where pedido in (" + var_pedidos_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
                  Me.lbl_hora_inicio = CStr(IIf(IsNull(rsaux7!HORA_INICIO), "", rsaux7!HORA_INICIO))
                  Me.lbl_hora_final = "PROCESO"
               End If
               rsaux7.Close
            Else
               rsaux7.Open "select MIN(HORA_INICIO) AS HORA_INICIO,  MAX(HORA_FINAL) as HORA_FINAL from TB_ORACLE_TIEMPO_POR_LOTE where pedido in(" + var_pedidos_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
                  Me.lbl_hora_inicio = CStr(IIf(IsNull(rsaux7!HORA_INICIO), "", rsaux7!HORA_INICIO))
                  Me.lbl_hora_final = CStr(IIf(IsNull(rsaux7!HORA_FINAL), "", rsaux7!HORA_FINAL))
               End If
               rsaux7.Close
            End If
            
            
            
            If var_cerrado = 0 Then
               rsaux7.Open "select  cast(getdate() - min(hora_inicio) as time) as tiempo from TB_ORACLE_TIEMPO_POR_LOTE where pedido in (" + var_pedidos_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
                  Me.lbl_tiempo_transcurrido = "Tiempo transcurrido:"
                  Me.lbl_tiempo = Mid(CStr(IIf(IsNull(rsaux7!tiempo), "", rsaux7!tiempo)), 1, 8)
               End If
               rsaux7.Close
            Else
               rsaux7.Open "select cast(max(hora_final) - min(hora_inicio) as time) as tiempo from TB_ORACLE_TIEMPO_POR_LOTE where pedido in(" + var_pedidos_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux7.EOF Then
               Me.lbl_tiempo_transcurrido = "Tiempo total:"
               'Me.lbl_tiempo = Mid(CStr(IIf(IsNull(rsaux7!tiempo), "", rsaux7!tiempo)), 1, 8)
               If Me.lbl_hora_final = "" Then
                  Me.lbl_hora_final = Now
               End If
               If Me.lbl_hora_inicio = "" Then
                  Me.lbl_hora_inicio = Now
               End If
               
               Me.lbl_tiempo = Mid(CDate(CDate(Me.lbl_hora_inicio) - CDate(Me.lbl_hora_final)), 1, 8)
               If Mid(Me.lbl_tiempo, 1, 2) = "12" Then
                  Me.lbl_tiempo = "00" + Mid(Me.lbl_tiempo, 3, Len(Me.lbl_tiempo))
               End If
               End If
               rsaux7.Close
            End If
            
            
            
            
            
            
            x = 0
            If x = 1 Then
            rsaux8.Open "select pedido, NUMERO_CAJA, estatus  from tb_oracle_cajas_aduana where ESTATUS in ('S','L') and embarque = " + Me.lbl_embarque, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  strconsulta = "select lote from xxvia_tb_salidas_cajas where source_header_number = ? and inte_paq_caja = ? and rownum = 1"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux8!pedido)
                       .Parameters.Append parametro
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux8!numero_caja)
                       .Parameters.Append parametro
                  End With
                  Set rsaux7 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux7.EOF Then
                     var_lote = rsaux7!lote
                     For var_j = 1 To Me.lv_pedidos.ListItems.Count
                         Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                         If CDbl(Me.lv_pedidos.selectedItem) = rsaux8!pedido And CDbl(Me.lv_pedidos.selectedItem.SubItems(1)) = rsaux7!lote Then
                            If rsaux8!estatus = "L" Then
                               Me.lv_pedidos.selectedItem.SubItems(11) = "Cargando"
                            End If
                            If rsaux8!estatus = "S" Then
                               Me.lv_pedidos.selectedItem.SubItems(11) = "Cerrado"
                            End If
                         End If
                     Next var_j
                  End If
                  rsaux7.Close
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            Else
            var_embarque_cerrado = 0
            If rsaux8.State = 1 Then
               rsaux8.Close
            End If
            rsaux8.Open "select distinct pedido, NUMERO_CAJA, estatus, lote  from tb_oracle_cajas_aduana where ESTATUS in ('S','L') and embarque = " + Me.lbl_embarque, cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux8.EOF
                  For var_j = 1 To Me.lv_pedidos.ListItems.Count
                      Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                      If CDbl(Me.lv_pedidos.selectedItem) = rsaux8!pedido And Me.lv_pedidos.selectedItem.SubItems(1) = rsaux8!lote Then
                         If rsaux8!estatus = "L" Then
                            Me.lv_pedidos.selectedItem.SubItems(11) = "Cargando"
                            var_embarque_cerrado = 1
                         End If
                         If rsaux8!estatus = "S" Then
                            Me.lv_pedidos.selectedItem.SubItems(11) = "Cerrado"
                         End If
                      End If
                  Next var_j
                  rsaux8.MoveNext
            Wend
            rsaux8.Close
            
            
            For var_j = 1 To Me.lv_pedidos.ListItems.Count
                Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                If Me.lv_pedidos.selectedItem.SubItems(11) = "" Then
                   var_embarque_cerrado = 1
                End If
            Next var_j
            
            
            
            End If
            var_fecha_fin_pantalla = Now
            Me.Label4 = CDate(CDate(var_fecha_fin_pantalla) - CDate(var_fecha_inicio_pantalla))
            Me.lv_pantalla.Visible = False
            rs.Open "SELECT Min(dbo.TB_ORACLE_TIEMPO_PEDIDO_ADUANAS.HORA_inicio) AS HORA_inicio, cast(getdate() - min(hora_inicio) as time) as tiempo  FROM  dbo.TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES INNER JOIN dbo.TB_ORACLE_TIEMPO_PEDIDO_ADUANAS ON dbo.TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES.PEDIDO = dbo.TB_ORACLE_TIEMPO_PEDIDO_ADUANAS.PEDIDO Where (dbo.TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES.Embarque = " + Me.lbl_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
            var_hora = CStr(IIf(IsNull(rs(0).Value), "", rs(0).Value))
            If Not rs.EOF Then
               If var_hora <> "" Then
                  Me.lbl_hora_inicio_aduana = CStr(rs(0).Value)
                  If var_embarque_cerrado = 0 Then
                     rsaux1.Open "SELECT Max(dbo.TB_ORACLE_TIEMPO_PEDIDO_ADUANAS.HORA_FIN) AS HORA_FIN FROM  dbo.TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES INNER JOIN dbo.TB_ORACLE_TIEMPO_PEDIDO_ADUANAS ON dbo.TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES.PEDIDO = dbo.TB_ORACLE_TIEMPO_PEDIDO_ADUANAS.PEDIDO Where (dbo.TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES.Embarque = " + Me.lbl_embarque + ")", cnn, adOpenDynamic, adLockOptimistic
                     Me.lbl_hora_final_aduana = CStr(rsaux1(0).Value)
                     Me.lbl_tiempo_transcurrido_aduana = "Tiempo total:"
                     Me.lbl_tiempo_aduana = Mid(CDate(CDate(Me.lbl_hora_inicio_aduana) - CDate(Me.lbl_hora_final_aduana)), 1, 8)
                     rsaux1.Close
                  Else
                     
                     Me.lbl_hora_final_aduana = "PROCESO"
                     Me.lbl_tiempo_transcurrido_aduana = "Tiempo transcurrido:"
                     Me.lbl_tiempo_aduana = Mid(CStr(IIf(IsNull(rs!tiempo), "", rs!tiempo)), 1, 8)
                  End If
               Else
                  Me.lbl_hora_inicio_aduana = ""
                  Me.lbl_hora_final_aduana = ""
                  Me.lbl_tiempo_aduana = ""
               End If
            End If
            rs.Close
            
            var_x = 1
            VAR_Y = 1
            cnn.BeginTrans
            rs.Open "SELECT MAX(INTE_tEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_PROGRESO_LOTES", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
            Else
               var_consecutivo = 0
            End If
            rs.Close
            var_consecutivo = var_consecutivo + 1
            rs.Open "INSERT INTO TB_TEMP_ORACLE_PROGRESO_LOTES (INTE_TEM_CONSECUTIVO) VALUES  (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_x = 1
            For var_j = 1 To Me.lv_pedidos.ListItems.Count
                Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                If var_x = 17 Then
                   If VAR_Y > 1 Then
                      Me.lv_pedidos.selectedItem.SubItems(12) = Me.lv_pedidos.selectedItem + CStr(VAR_Y)
                      rs.Open "INSERT INTO TB_TEMP_ORACLE_PROGRESO_LOTES (INTE_TEM_CONSECUTIVO, PEDIDO) VALUES  (" + CStr(var_consecutivo) + ",'" + Me.lv_pedidos.selectedItem + CStr(VAR_Y) + "')", cnn, adOpenDynamic, adLockOptimistic
                   End If
                   
                   VAR_Y = VAR_Y + 1
                   var_x = 0
                End If
                var_x = var_x + 1
            
            Next var_j
            If Me.lv_pedidos.ListItems.Count > 0 Then
               Me.lv_pedidos.selectedItem.SubItems(12) = Me.lv_pedidos.selectedItem + CStr(VAR_Y)
               rs.Open "INSERT INTO TB_TEMP_ORACLE_PROGRESO_LOTES (INTE_TEM_CONSECUTIVO, PEDIDO) VALUES  (" + CStr(var_consecutivo) + ",'" + Me.lv_pedidos.selectedItem + CStr(VAR_Y) + "')", cnn, adOpenDynamic, adLockOptimistic
               rs.Open "DELETE FROM TB_TEMP_ORACLE_PROGRESO_LOTES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND PEDIDO IS NULL", cnn, adOpenDynamic, adLockOptimistic
               var_veces = VAR_Y
               If var_veces > 1 Then
                  z = 1
                  If z = 0 Then
                     For var_j = 18 To Me.lv_pedidos.ListItems.Count
                          'Sleep 1000
                          lv_pedidos.ListItems.Item(var_j).Selected = True
                          lv_pedidos.selectedItem.EnsureVisible
                          Me.Refresh
                      Next var_j
                  Else
                     'Sleep 6000
                     rs.Open "SELECT * FROM TB_TEMP_ORACLE_PROGRESO_LOTES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                     While Not rs.EOF
                           For var_j = 1 To Me.lv_pedidos.ListItems.Count
                               Me.lv_pedidos.ListItems.Item(var_j).Selected = True
                               If Me.lv_pedidos.selectedItem.SubItems(12) = rs!pedido Then
                                  lv_pedidos.selectedItem.EnsureVisible
                                  Me.Refresh
                                  'Sleep 6000
                               End If
                           Next var_j
                           rs.MoveNext
                     Wend
                     rs.Close
                  End If
               End If
           End If
           Me.Timer1.Enabled = True
           If Me.lbl_hora_inicio = Me.lbl_hora_final Then
              Me.lbl_hora_inicio = ""
              Me.lbl_hora_final = ""
           End If
           rs.Open "DELETE FROM TB_TEMP_ORACLE_PROGRESO_LOTES WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   
   
        lv_pedidos.Sorted = True
        lv_pedidos.SortKey = 0
        lv_pedidos.SortOrder = lvwAscending
   
   
        lv_pedidos.Sorted = True
        lv_pedidos.SortKey = 2
        lv_pedidos.SortOrder = lvwAscending
   
   
   
   Me.frm_procesando.Visible = False
End Sub

Private Sub Form_Activate()
   Me.Command1.SetFocus
End Sub

Private Sub Form_Load()
   Me.frm_procesando.Visible = False
   
   Me.Timer2.Enabled = False
   'Me.lv_pantalla.Visible = True
   var_fecha_inicio_pantalla = Now
   Me.lbl_tiempo = ""
   Me.lbl_tiempo_transcurrido = ""
   Me.lbl_surtir = "0.00"
   Me.lbl_cantidad_surtida = "0.00"
   Me.lbl_porcentaje = "0.00%"
   Me.lbl_embarque = Format(var_embarque_lotes, "#########")

            
End Sub

Private Sub lv_pantalla_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_pedidos, ColumnHeader)
End Sub

Private Sub lv_pantalla_GotFocus()
   Me.Timer1.Enabled = False
End Sub

Private Sub lv_pedidos_GotFocus()
   Me.Timer1.Enabled = False
End Sub

Private Sub lv_pedidos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub lv_pedidos_LostFocus()
   Me.Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   'Call Command1_Click
End Sub

