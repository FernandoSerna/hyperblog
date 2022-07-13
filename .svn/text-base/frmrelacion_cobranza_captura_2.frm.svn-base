VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrelacion_cobranza_captura_2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura de Cobranza"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cerrar_relacion 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1455
      Picture         =   "frmrelacion_cobranza_captura_2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Cerrar relación de cobranza"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cancelar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11235
      Picture         =   "frmrelacion_cobranza_captura_2.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   795
      Picture         =   "frmrelacion_cobranza_captura_2.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   135
      Picture         =   "frmrelacion_cobranza_captura_2.frx":0886
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1125
      Picture         =   "frmrelacion_cobranza_captura_2.frx":0988
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      Picture         =   "frmrelacion_cobranza_captura_2.frx":0AD2
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Buscar folio"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_busqueda 
      Height          =   960
      Left            =   480
      TabIndex        =   49
      Top             =   450
      Width           =   3135
      Begin VB.TextBox txt_busqueda_folio 
         Height          =   315
         Left            =   195
         MaxLength       =   10
         TabIndex        =   50
         Top             =   510
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   " Folio de la relación"
         ForeColor       =   &H8000000E&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   51
         Top             =   120
         Width           =   3075
      End
   End
   Begin VB.TextBox txt_aplicada 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   46
      Top             =   15
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Frame Frame8 
      Height          =   4515
      Left            =   6270
      TabIndex        =   45
      Top             =   420
      Width           =   5310
      Begin VB.TextBox txt_total_relacion 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   2865
         TabIndex        =   53
         Top             =   4005
         Width           =   2370
      End
      Begin MSComctlLib.ListView lv_relacion 
         Height          =   3765
         Left            =   60
         TabIndex        =   61
         Top             =   165
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   6641
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
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "RELACION"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "FECHA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "AGENTE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "NOMBRE AGENTE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CHEQUE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "BANCO CHEQUE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "NOMBRE BANCO CHEQUE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FECHA CHEQUE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Documento"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Serie"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Número"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "AGENTE DOCUMENTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "NOMBRE AGENTE DOCUMENTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "CLIENTE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "NOMBRE CLIENTE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "FECHA "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "IMPORTE DOCUMENTO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "SALDO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "%"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "aplicada"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         TabIndex        =   52
         Top             =   4020
         Width           =   795
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   300
      TabIndex        =   39
      Top             =   420
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   62
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
         TabIndex        =   40
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   " Cheque "
      Height          =   45
      Left            =   150
      TabIndex        =   41
      Top             =   1650
      Visible         =   0   'False
      Width           =   6060
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Picture         =   "frmrelacion_cobranza_captura_2.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Nuevo Movimiento"
         Top             =   240
         Width           =   330
      End
      Begin VB.Frame Frame6 
         Height          =   45
         Left            =   30
         TabIndex        =   47
         Top             =   600
         Width           =   5985
      End
      Begin VB.TextBox txt_fecha_cheque 
         Height          =   345
         Left            =   4890
         TabIndex        =   7
         Top             =   675
         Width           =   990
      End
      Begin VB.TextBox txt_cheque 
         Height          =   345
         Left            =   780
         MaxLength       =   4
         TabIndex        =   4
         Top             =   690
         Width           =   720
      End
      Begin VB.TextBox txt_nombre_banco_cheque 
         Height          =   345
         Left            =   2685
         TabIndex        =   6
         Top             =   690
         Width           =   1590
      End
      Begin VB.TextBox txt_banco_cheque 
         Height          =   345
         Left            =   2130
         TabIndex        =   5
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
         Height          =   195
         Left            =   90
         TabIndex        =   44
         Top             =   750
         Width           =   600
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4365
         TabIndex        =   43
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   1590
         TabIndex        =   42
         Top             =   750
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos Relación "
      Height          =   1185
      Left            =   150
      TabIndex        =   35
      Top             =   435
      Width           =   6060
      Begin VB.TextBox txt_relacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   0
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txt_agente 
         Height          =   330
         Left            =   1425
         TabIndex        =   2
         Top             =   720
         Width           =   750
      End
      Begin VB.TextBox txt_nombre_agente 
         Height          =   330
         Left            =   2190
         TabIndex        =   3
         Top             =   720
         Width           =   3675
      End
      Begin VB.TextBox txt_fecha 
         Height          =   345
         Left            =   3810
         TabIndex        =   1
         Top             =   315
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Relación:"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   345
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   37
         Top             =   735
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3225
         TabIndex        =   36
         Top             =   375
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documentos a aplicar "
      Height          =   1125
      Left            =   150
      TabIndex        =   31
      Top             =   1605
      Width           =   6060
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Picture         =   "frmrelacion_cobranza_captura_2.frx":0CD6
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Nuevo Movimiento"
         Top             =   210
         Width           =   330
      End
      Begin VB.Frame Frame9 
         Height          =   45
         Left            =   30
         TabIndex        =   48
         Top             =   540
         Width           =   5985
      End
      Begin VB.TextBox txt_documento 
         Height          =   315
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   8
         Top             =   735
         Width           =   1020
      End
      Begin VB.TextBox txt_numero 
         Height          =   315
         Left            =   4575
         TabIndex        =   10
         Top             =   735
         Width           =   1320
      End
      Begin VB.TextBox txt_serie 
         Height          =   315
         Left            =   3060
         TabIndex        =   9
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   795
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   3840
         TabIndex        =   33
         Top             =   795
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   2520
         TabIndex        =   32
         Top             =   795
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Datos del Documento "
      Height          =   1455
      Left            =   150
      TabIndex        =   25
      Top             =   2745
      Width           =   6060
      Begin VB.TextBox txt_agente_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   780
         TabIndex        =   11
         Top             =   300
         Width           =   1260
      End
      Begin VB.TextBox txt_nombre_agente_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2070
         TabIndex        =   12
         Top             =   300
         Width           =   3855
      End
      Begin VB.TextBox txt_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   780
         TabIndex        =   13
         Top             =   645
         Width           =   1260
      End
      Begin VB.TextBox txt_nombre_cliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2070
         TabIndex        =   14
         Top             =   645
         Width           =   3855
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   2460
         TabIndex        =   16
         Top             =   975
         Width           =   1335
      End
      Begin VB.TextBox txt_fecha_factura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   780
         TabIndex        =   15
         Top             =   990
         Width           =   990
      End
      Begin VB.TextBox txt_saldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   4485
         TabIndex        =   17
         Top             =   975
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   375
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   135
         TabIndex        =   29
         Top             =   705
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   1875
         TabIndex        =   28
         Top             =   1005
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   195
         Left            =   3945
         TabIndex        =   26
         Top             =   1065
         Width           =   450
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Importe a aplicar "
      Height          =   720
      Left            =   150
      TabIndex        =   21
      Top             =   4215
      Width           =   6060
      Begin VB.TextBox txt_importe_aplicar 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1095
         TabIndex        =   18
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txt_descuento 
         Height          =   315
         Left            =   4395
         MaxLength       =   1
         TabIndex        =   19
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
         Height          =   195
         Left            =   405
         TabIndex        =   24
         Top             =   330
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Left            =   3285
         TabIndex        =   23
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   5160
         TabIndex        =   22
         Top             =   330
         Width           =   120
      End
   End
   Begin VB.Frame Frame5 
      Height          =   75
      Left            =   90
      TabIndex        =   20
      Top             =   300
      Width           =   11565
   End
End
Attribute VB_Name = "frmrelacion_cobranza_captura_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tipo_lista As Integer
Dim var_total_relacion As Double
Dim var_primera_vez As Integer
Dim var_ventana As Integer
Private Sub cmd_aceptar_pedidos_Click()
    If Trim(txt_agente) <> "" Then
       If IsDate(Me.txt_fecha) Then
          If Trim(txt_documento) <> "" Then
             If Trim(txt_numero) <> "" Then
                If Trim(txt_cliente) <> "" Then
                   If IsNumeric(Me.txt_importe_aplicar) Then
                      If Trim(Me.txt_descuento) = "" Then
                         Me.txt_descuento = "0"
                      End If
                      If IsNumeric(txt_descuento) Then
                         'If Trim(txt_cheque) <> "" Then
                            'If Trim(txt_banco_cheque) <> "" Then
                               'If IsDate(Me.txt_fecha_cheque) Then
                                  If Me.txt_aplicada <> "*" Then
                                     rsaux9.Open "SELECT * FROM TB_RELACION_COBRANZA with (nolock) WHERE VCHA_rCO_FOLIO = '" + Me.txt_relacion + "' AND VCHA_aGE_AGENTE_ID = '" + Me.txt_agente + "' AND VCHA_rCO_CHEQUE = '" + Me.txt_cheque + "' AND VCHA_BAN_BANCO_ID = '" + Me.txt_banco_cheque + "' AND VCHA_cAR_DOCUMENTO = '" + Me.txt_documento + "' AND INTE_cAR_NUMERO = " + Me.txt_numero + " AND VCHA_sER_SERIE_ID = '" + Me.txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                                     If rsaux9.EOF Then
                                        var_si = MsgBox("¿Desea aplicar la cobranza?", vbYesNo, "ATENCION")
                                        If var_si = 6 Then
                                           var_si = MsgBox("Confirmar la aplicación del pago", vbYesNo, "ATENCION")
                                           If var_si = 6 Then
                                              If txt_agente <> Me.txt_agente_factura Then
                                                 var_si = 0
                                                 var_si = MsgBox("la factura no corresponde al agente seleccionado, ¿Desea aplicar el pago?", vbYesNo, "ATENCION")
                                              End If
                                              If var_si = 6 Then
                                                 rs.Open "SELECT MAX(INTE_RCO_PARTIDA) FROM TB_RELACION_COBRANZA WHERE VCHA_RCO_FOLIO = '" + txt_relacion + "'", cnn, adOpenDynamic, adLockOptimistic
                                                 If Not rs.EOF Then
                                                    var_partida = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                                                 Else
                                                   var_partida = 1
                                                 End If
                                                 rs.Close
                                                 
                                                 If var_primera_vez = 1 Then
                                                    var_primera_vez = 0
                                                    rs.Open "select max(cast(vcha_rco_folio as bigint)) from tb_relacion_cobranza where len(vcha_rco_folio) = 10 and substring(vcha_rco_folio,1,5) = '00000'", cnn, adOpenDynamic, adLockOptimistic
                                                    If rs.EOF Then
                                                       var_consecutivo = 0
                                                    Else
                                                       var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
                                                    End If
                                                    rs.Close
                                                    
                                                    If var_consecutivo < 10000 Then
                                                       var_consecutivo = 10000
                                                    End If
                                                    var_consecutivo = var_consecutivo + 1
                                                    var_consecutivo_str = CStr(var_consecutivo)
                                                    var_j = Len(var_consecutivo_str)
                                                    For var_i = var_j + 1 To 10
                                                        var_consecutivo_str = "0" + var_consecutivo_str
                                                    Next var_i
                                                    Me.txt_relacion = var_consecutivo_str
                                                 End If
                                                 var_primera_vez = 0
                                                 
                                                 
                                                 var_dia = CStr(Day(Date))
                                                 var_mes = CStr(Month(Date))
                                                 var_año = CStr(Year(Date))
                                                 If Len(Trim(var_dia)) = 1 Then
                                                    var_dia = "0" + var_dia
                                                 End If
                                                 If Len(Trim(var_mes)) = 1 Then
                                                    var_mes = "0" + var_mes
                                                 End If
                                                 var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                                                 
                                                 
                                                 
                                                 Cadena = "EXECUTE RELACION_COBRANZA_DEPOSITO '" + var_empresa + "', '" + var_unidad_organizacional + "', '" + Me.txt_relacion + "', '" + Me.txt_fecha + "', '" + Me.txt_agente + "', '" + Me.txt_cliente + "', '0001', '" + Me.txt_fecha + "', " + Me.txt_importe_aplicar + ", " + Me.txt_descuento + ", " + txt_numero + ", 0, 0, " + CStr(var_partida) + ", 0, '" + txt_serie + "', '" + txt_documento + "', 'EFVO', 'SIN REFERENCIA', '" + Me.txt_fecha + "', 'EFVO',0"
                                                 rs.Open Cadena, cnn, adOpenDynamic, adLockOptimistic
                                                 rs.Open "update tb_relacion_cobranza set dtim_rco_fecha_insercion = getdate(), vcha_rco_estatus = '' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion + "' and inte_rco_partida = " + CStr(var_partida), cnn, adOpenDynamic, adLockOptimistic
                                                 Set list_item = lv_relacion.ListItems.Add(, , Me.txt_relacion)
                                                 list_item.SubItems(1) = Me.txt_fecha
                                                 list_item.SubItems(2) = Me.txt_agente
                                                 list_item.SubItems(3) = Me.txt_nombre_agente
                                                 list_item.SubItems(4) = Me.txt_cheque
                                                 list_item.SubItems(5) = Me.txt_banco_cheque
                                                 list_item.SubItems(6) = Me.txt_nombre_banco_cheque
                                                 list_item.SubItems(7) = Me.txt_fecha_cheque
                                                 list_item.SubItems(8) = Me.txt_documento
                                                 list_item.SubItems(9) = Me.txt_serie
                                                 list_item.SubItems(10) = Me.txt_numero
                                                 list_item.SubItems(11) = Me.txt_agente_factura
                                                 list_item.SubItems(12) = Me.txt_nombre_agente_factura
                                                 list_item.SubItems(13) = Me.txt_cliente
                                                 list_item.SubItems(14) = Me.txt_nombre_cliente
                                                 list_item.SubItems(15) = Me.txt_fecha_factura
                                                 list_item.SubItems(16) = Me.txt_importe
                                                 list_item.SubItems(17) = Me.txt_saldo
                                                 list_item.SubItems(18) = Me.txt_importe_aplicar
                                                 list_item.SubItems(19) = Me.txt_descuento
                                                 list_item.SubItems(20) = ""
                                                 Me.txt_total_relacion = Format(CDbl(Me.txt_total_relacion) + CDbl(Me.txt_importe_aplicar), "###,###,##0.00")
                                                 MsgBox "Se a cargado la relación correctamente", vbOKOnly, "ATENCION"
                                              End If
                                           End If
                                        End If
                                     Else
                                         MsgBox "El pago ya existe en la relación", vbOKOnly, "ATENCION"
                                     End If
                                     rsaux9.Close
                                  Else
                                     MsgBox "La relación ya fue aplicada", vbOKOnly, "ATENCION"
                                  End If
                         '      Else
                         '         MsgBox "Fecha de cheque invalida", vbOKOnly, "ATENCION"
                         '      End If
                         '   Else
                         '      MsgBox "Debe de indicar un banco para el cheque", vbOKOnly, "ATENCION"
                         '   End If
                         'Else
                         '   MsgBox "Debe de indicar un cheque", vbOKOnly, "ATENCION"
                         'End If
                      Else
                         MsgBox "Descuento invalido", vbOKOnly, "ATENCION"
                      End If
                   Else
                      MsgBox "Importe a aplicar incorrecto"
                   End If
                Else
                   MsgBox "Clave de cliente incorrecta", vbOKOnly, "TENCION"
                End If
             Else
                MsgBox "Numero de documento incorrecto", vbOKOnly, "ATENCION"
             End If
          Else
             MsgBox "No se a seleccionado un tipo de documento", vbOKOnly, "ATENCION"
          End If
       Else
          MsgBox "Fecha de relación de cobranza incorrecta", vbOKOnly, "ATENCION"
       End If
    Else
       MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
    End If
End Sub

Private Sub cmd_buscar_Click()
   var_primera_vez = 0
   Me.txt_busqueda_folio = ""
   Me.frm_busqueda.Visible = True
   var_ventana = 1
   Me.txt_busqueda_folio.SetFocus
End Sub

Private Sub cmd_cancelar_pedidos_Click()
   Unload Me
End Sub

Private Sub cmd_cerrar_relacion_Click()
   If Me.txt_relacion <> "" Then
      var_si = MsgBox("¿Desea cerrar la relación?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar el cerrado de la relación", vbYesNo, "ATENCION")
         If var_si = 6 Then
            rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_rco_folio = '" + Me.txt_relacion + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux.Open "update tb_relacion_cobranza set vcha_rco_estatus = 'I' where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_rco_folio = '" + txt_relacion + "' and inte_rco_partida = " + CStr(rs!inte_rco_partida), cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            Me.cmd_aceptar_pedidos.Enabled = False
            Me.cmd_eliminar.Enabled = False
            MsgBox "Se a cerrado la relación de corbanza", vbOKOnly, "ATENCION"
          End If
      End If
   Else
      MsgBox "No se a seleccionado una relación de cobranza", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_eliminar_Click()
   If Me.txt_relacion <> "" Then
      If Me.txt_agente <> "" Then
         If Me.txt_cheque <> "" Then
            If Me.txt_banco_cheque <> "" Then
               If Me.txt_documento <> "" Then
                  If IsNumeric(Me.txt_numero) Then
                     If Trim(Me.txt_aplicada) <> "*" Then
                        var_si = MsgBox("¿Desea eliminar el registro de la relación de cobranza?", vbYesNo, "ATENCIO")
                        If var_si = 6 Then
                           var_si = MsgBox("Confirmar la eliminación del registro de la relación de cobranza", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              rsaux9.Open "DELETE FROM TB_RELACION_COBRANZA WHERE VCHA_rCO_FOLIO = '" + Me.txt_relacion + "' AND VCHA_aGE_AGENTE_ID = '" + Me.txt_agente + "' AND VCHA_rCO_CHEQUE = '" + Me.txt_cheque + "' AND VCHA_BAN_BANCO_ID = '" + Me.txt_banco_cheque + "' AND VCHA_cAR_DOCUMENTO = '" + Me.txt_documento + "' AND INTE_cAR_NUMERO = " + Me.txt_numero + " AND VCHA_sER_SERIE_ID = '" + Me.txt_serie + "'", cnn, adOpenDynamic, adLockOptimistic
                              Me.txt_total_relacion = Format(CDbl(Me.txt_total_relacion) - CDbl(Me.lv_relacion.selectedItem.SubItems(18)), "###,###,##0.00")
                              lv_relacion.ListItems.Remove (lv_relacion.selectedItem.Index)
                              If lv_relacion.ListItems.Count > 0 Then
                                 Me.lv_relacion.SetFocus
                              Else
                                 Me.txt_agente = ""
                                 Me.txt_agente_factura = ""
                                 Me.txt_cliente = ""
                                 Me.txt_descuento = ""
                                 Me.txt_documento = ""
                                 Me.txt_fecha = Date
                                 Me.txt_fecha_factura = Date
                                 Me.txt_importe = ""
                                 Me.txt_importe_aplicar = ""
                                 Me.txt_nombre_agente = ""
                                 Me.txt_nombre_agente_factura = ""
                                 Me.txt_nombre_cliente = ""
                                 Me.txt_numero = ""
                                 Me.txt_relacion = ""
                                 Me.txt_saldo = ""
                                 Me.txt_serie = ""
                                 Me.txt_cheque = ""
                                 Me.txt_banco_cheque = ""
                                 Me.txt_nombre_banco_cheque = ""
                                 'Me.txt_relacion.SetFocus
                              End If
                           End If
                        End If
                     Else
                        MsgBox "La relación ya no puede ser eliminada", vbOKOnly, "ATENCION"
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.cmd_aceptar_pedidos.Enabled = True
   Me.cmd_eliminar.Enabled = True
   Me.txt_relacion = ""
   Me.lv_relacion.ListItems.Clear
   Me.txt_agente = ""
   Me.txt_agente_factura = ""
   Me.txt_cliente = ""
   Me.txt_descuento = ""
   Me.txt_documento = ""
   Me.txt_fecha = Date
   Me.txt_fecha_factura = Date
   Me.txt_importe = ""
   Me.txt_importe_aplicar = ""
   Me.txt_nombre_agente = ""
   Me.txt_nombre_agente_factura = ""
   Me.txt_nombre_cliente = ""
   Me.txt_numero = ""
   Me.txt_saldo = ""
   Me.txt_serie = ""
   Me.txt_cheque = ""
   Me.txt_banco_cheque = ""
   Me.txt_nombre_banco_cheque = ""
   Me.txt_fecha.SetFocus
End Sub

Private Sub Command1_Click()
   If Me.txt_relacion = "" Then
      MsgBox "No a indicado un folio de relación", vbOKOnly, "ATENCION"
   Else
      Me.txt_fecha_cheque = Date
      Me.txt_agente_factura = ""
      Me.txt_nombre_agente_factura = ""
      Me.txt_cheque = ""
      Me.txt_banco_cheque = ""
      Me.txt_nombre_banco_cheque = ""
      Me.txt_cliente = ""
      Me.txt_nombre_cliente = ""
      Me.txt_fecha_factura = Date
      Me.txt_importe = ""
      Me.txt_saldo = ""
      Me.txt_documento = ""
      Me.txt_serie = ""
      Me.txt_numero = ""
      Me.txt_importe_aplicar = ""
      Me.txt_descuento = ""
      Me.txt_aplicada = ""
      Me.txt_banco_cheque.SetFocus
   End If
End Sub

Private Sub Command2_Click()
      Me.txt_agente_factura = ""
      Me.txt_nombre_agente_factura = ""
      Me.txt_cliente = ""
      Me.txt_nombre_cliente = ""
      Me.txt_fecha_factura = Date
      Me.txt_importe = ""
      Me.txt_saldo = ""
      Me.txt_documento = ""
      Me.txt_serie = ""
      Me.txt_numero = ""
      Me.txt_importe_aplicar = ""
      Me.txt_descuento = ""
      Me.txt_aplicada = ""
      Me.txt_documento = "FA"
      rs.Open "select * from tb_series where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_serie = IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID)
      Else
         Me.txt_serie = ""
      End If
      rs.Close
      Me.txt_documento.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 66 Then
      cmd_buscar_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If var_ventana = 0 Then
         Unload Me
      Else
         If var_ventana = 1 Then
            var_ventana = 0
            Me.frm_busqueda.Visible = False
         Else
            If var_ventana = 2 Then
               var_ventana = 0
               Me.frm_lista.Visible = False
            End If
         End If
      End If
   End If
End Sub

Private Sub Form_Load()
   var_primera_vez = 1
   Top = 1500
   Left = 0
   Me.txt_relacion = ""
   frm_lista.Visible = False
   Me.txt_fecha = Date
   Me.txt_fecha_cheque = Date
   Me.txt_relacion.Enabled = False
   Me.frm_busqueda.Visible = False
   Me.txt_total_relacion = "0"
   var_ventana = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         If var_tipo_lista = 3 Then
            txt_agente = lv_lista.selectedItem
            txt_nombre_agente = lv_lista.selectedItem.SubItems(1)
         End If
         If var_tipo_lista = 1 Then
            txt_banco = lv_lista.selectedItem
            txt_nombre_banco = lv_lista.selectedItem.SubItems(1)
         End If
         If var_tipo_lista = 2 Then
            txt_banco_cheque = lv_lista.selectedItem
            txt_nombre_banco_cheque = lv_lista.selectedItem.SubItems(1)
         End If
         If var_tipo_lista = 4 Then
            Me.txt_cliente = lv_lista.selectedItem
            Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
            Me.txt_agente_factura = Me.txt_agente
            Me.txt_nombre_agente_factura = Me.txt_nombre_agente
         End If
      Else
         txt_agente = ""
         txt_nombre_agente = ""
      End If
      If var_tipo_lista = 3 Then
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 1 Then
         txt_banco.SetFocus
      End If
      If var_tipo_lista = 2 Then
         txt_banco_cheque.SetFocus
      End If
      If var_tipo_lista = 4 Then
         Me.txt_importe_aplicar.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      If var_tipo_lista = 3 Then
         txt_agente.SetFocus
      End If
      If var_tipo_lista = 1 Then
         txt_banco.SetFocus
      End If
      If var_tipo_lista = 2 Then
         txt_banco_cheque.SetFocus
      End If
      If var_tipo_lista = 4 Then
         Me.txt_documento = ""
         Me.txt_documento.SetFocus
      End If
   End If
End Sub

Private Sub lv_lista_LostFocus()
   var_ventana = 0
   frm_lista.Visible = False
End Sub

Private Sub lv_relacion_GotFocus()
   If Me.lv_relacion.ListItems.Count > 0 Then
      Me.txt_relacion = Me.lv_relacion.selectedItem
      Me.txt_fecha = Me.lv_relacion.selectedItem.SubItems(1)
      Me.txt_agente = Me.lv_relacion.selectedItem.SubItems(2)
      Me.txt_nombre_agente = Me.lv_relacion.selectedItem.SubItems(3)
      Me.txt_cheque = Me.lv_relacion.selectedItem.SubItems(4)
      Me.txt_banco_cheque = Me.lv_relacion.selectedItem.SubItems(5)
      Me.txt_nombre_banco_cheque = Me.lv_relacion.selectedItem.SubItems(6)
      Me.txt_fecha_cheque = Me.lv_relacion.selectedItem.SubItems(7)
      Me.txt_documento = Me.lv_relacion.selectedItem.SubItems(8)
      Me.txt_serie = Me.lv_relacion.selectedItem.SubItems(9)
      Me.txt_numero = Me.lv_relacion.selectedItem.SubItems(10)
      Me.txt_agente_factura = Me.lv_relacion.selectedItem.SubItems(11)
      Me.txt_nombre_agente_factura = Me.lv_relacion.selectedItem.SubItems(12)
      Me.txt_cliente = Me.lv_relacion.selectedItem.SubItems(13)
      Me.txt_nombre_cliente = Me.lv_relacion.selectedItem.SubItems(14)
      Me.txt_fecha_factura = Me.lv_relacion.selectedItem.SubItems(15)
      Me.txt_importe = Format(Me.lv_relacion.selectedItem.SubItems(16), "###,###,##0.00")
      Me.txt_saldo = Format(Me.lv_relacion.selectedItem.SubItems(17), "###,###,##0.00")
      Me.txt_importe_aplicar = Format(Me.lv_relacion.selectedItem.SubItems(18), "###,###,##0.00")
      Me.txt_descuento = Me.lv_relacion.selectedItem.SubItems(19)
      Me.txt_aplicada = Me.lv_relacion.selectedItem.SubItems(20)
   End If
End Sub

Private Sub lv_relacion_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Me.txt_relacion = Me.lv_relacion.selectedItem
   Me.txt_fecha = Me.lv_relacion.selectedItem.SubItems(1)
   Me.txt_agente = Me.lv_relacion.selectedItem.SubItems(2)
   Me.txt_nombre_agente = Me.lv_relacion.selectedItem.SubItems(3)
   Me.txt_cheque = Me.lv_relacion.selectedItem.SubItems(4)
   Me.txt_banco_cheque = Me.lv_relacion.selectedItem.SubItems(5)
   Me.txt_nombre_banco_cheque = Me.lv_relacion.selectedItem.SubItems(6)
   Me.txt_fecha_cheque = Me.lv_relacion.selectedItem.SubItems(7)
   Me.txt_documento = Me.lv_relacion.selectedItem.SubItems(8)
   Me.txt_serie = Me.lv_relacion.selectedItem.SubItems(9)
   Me.txt_numero = Me.lv_relacion.selectedItem.SubItems(10)
   Me.txt_agente_factura = Me.lv_relacion.selectedItem.SubItems(11)
   Me.txt_nombre_agente_factura = Me.lv_relacion.selectedItem.SubItems(12)
   Me.txt_cliente = Me.lv_relacion.selectedItem.SubItems(13)
   Me.txt_nombre_cliente = Me.lv_relacion.selectedItem.SubItems(14)
   Me.txt_fecha_factura = Me.lv_relacion.selectedItem.SubItems(15)
   Me.txt_importe = Format(Me.lv_relacion.selectedItem.SubItems(16), "###,###,##0.00")
   Me.txt_saldo = Format(Me.lv_relacion.selectedItem.SubItems(17), "###,###,##0.00")
   Me.txt_importe_aplicar = Format(Me.lv_relacion.selectedItem.SubItems(18), "###,###,##0.00")
   Me.txt_descuento = Me.lv_relacion.selectedItem.SubItems(19)
   Me.txt_aplicada = Me.lv_relacion.selectedItem.SubItems(20)
   
   
End Sub

Private Sub txt_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES where vcha_emp_empresa_id = '" + var_empresa + "' order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 3
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
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_agente_LostFocus()
   If Trim(txt_agente) <> "" Then
      rs.Open "SELECT * FROM TB_aGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_agente = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
      Else
         MsgBox "Clave de agente incorrecta", vbOKOnly, "ATENCION"
         txt_agente = ""
         txt_nombre_agente = ""
      End If
      rs.Close
   Else
      txt_nombre_agente = ""
   End If
End Sub

Private Sub txt_banco_cheque_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL CHEQUE"
      var_tipo_lista = 2
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

Private Sub txt_banco_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_banco_cheque_LostFocus()
   If Trim(txt_banco_cheque) <> "" Then
      rs.Open "SELECT * FROM TB_BANCOS WHERE VCHA_BAN_BANCO_ID = '" + Me.txt_banco_cheque + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_banco_cheque = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
      Else
         MsgBox "Clave de banco incorrecto", vbOKOnly, "ATENCION"
         Me.txt_banco_cheque = ""
         Me.txt_nombre_banco_cheque = ""
      End If
      rs.Close
   End If
End Sub

Private Sub txt_banco_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos where vcha_ban_banco_id = '20' or vcha_ban_banco_id = '11' or vcha_ban_banco_id = '10' or vcha_ban_banco_id = '22' order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL DEPOSITO"
      var_tipo_lista = 1
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

Private Sub txt_banco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub


Private Sub txt_busqueda_folio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Me.txt_busqueda_folio) <> "" Then
         rs.Open "select *, isnull(vcha_Rco_Estatus,'') as estatus from tb_relacion_cobranza where vcha_Rco_folio = '" + txt_busqueda_folio + "' and vcha_emp_empresa_id = '" + var_empresa + "'", cnn, adOpenDynamic, adLockOptimistic
         Me.lv_relacion.ListItems.Clear
         If Not rs.EOF Then
            var_primera_vez = 0
            Me.txt_total_relacion = "0"
            While Not rs.EOF
                  Set list_item = lv_relacion.ListItems.Add(, , rs!vcha_Rco_folio)
                  list_item.SubItems(1) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
                  list_item.SubItems(2) = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                  rsaux.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     list_item.SubItems(3) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
                  Else
                     list_item.SubItems(3) = ""
                  End If
                  rsaux.Close
                  list_item.SubItems(4) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                  list_item.SubItems(5) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                  rsaux.Open "SELECT * FROM TB_BANCOS WHERE VCHA_BAN_BANCO_ID = '" + IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     list_item.SubItems(6) = IIf(IsNull(rsaux!VCHA_BAN_NOMBRE), "", rsaux!VCHA_BAN_NOMBRE)
                  Else
                     list_item.SubItems(6) = ""
                  End If
                  rsaux.Close
                  list_item.SubItems(7) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                  list_item.SubItems(8) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                  list_item.SubItems(9) = IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID)
                  list_item.SubItems(10) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
                  rsaux.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento = '" + IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento) + "' and vcha_Ser_Serie_id = '" + IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID) + "' and inte_Car_numero  = " + CStr(IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)), cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_agente = IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID)
                     var_cliente = IIf(IsNull(rsaux!vcha_cli_clave_id), "", rsaux!vcha_cli_clave_id)
                     list_item.SubItems(11) = var_agente
                     rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + var_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        list_item.SubItems(12) = IIf(IsNull(rsaux2!VCHA_AGE_NOMBRE), "", rsaux2!VCHA_AGE_NOMBRE)
                     Else
                        list_item.SubItems(12) = ""
                     End If
                     rsaux2.Close
                     rsaux2.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                     list_item.SubItems(13) = var_cliente
                     If Not rsaux2.EOF Then
                        list_item.SubItems(14) = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
                     Else
                        list_item.SubItems(14) = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
                     End If
                     rsaux2.Close
                     list_item.SubItems(15) = IIf(IsNull(rsaux!dtim_Car_fecha), "", rsaux!dtim_Car_fecha)
                     list_item.SubItems(16) = IIf(IsNull(rsaux!floa_Car_importe_neto), "", rsaux!floa_Car_importe_neto)
                     rsaux2.Open "select * from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento) + "' and vcha_ser_serie_id = '" + IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID) + "' and inte_Car_numero  = " + CStr(IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)), cnn, adOpenDynamic, adLockOptimistic
                     If Not rsaux2.EOF Then
                        list_item.SubItems(17) = IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE)
                     Else
                        list_item.SubItems(17) = 0
                     End If
                     rsaux2.Close
                     
                  Else
                  End If
                  rsaux.Close
                  list_item.SubItems(18) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,###,##0.00")
                  Me.txt_total_relacion = Format(CDbl(Me.txt_total_relacion) + IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,###,##0.00")
                  list_item.SubItems(19) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                  list_item.SubItems(20) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                  rs.MoveNext
            Wend
            If lv_relacion.ListItems.Count > 0 Then
               lv_relacion.ListItems.Item(1).Selected = True
               Me.lv_relacion.SetFocus
            End If
            rs.MoveFirst
            If rs!Estatus = "" Then
               Me.cmd_aceptar_pedidos.Enabled = True
               Me.cmd_eliminar.Enabled = True
            Else
               MsgBox "La relación de cobranza ya no puede ser modificada", vbOKOnly, "ATENCION"
               Me.cmd_aceptar_pedidos.Enabled = False
               Me.cmd_eliminar.Enabled = False
            End If
         Else
            MsgBox "La relación de cobranza no existe", vbOKOnly, "ATENCION "
            Me.txt_agente = ""
            Me.txt_agente_factura = ""
            Me.txt_cliente = ""
            Me.txt_descuento = ""
            Me.txt_documento = ""
            Me.txt_fecha = Date
            Me.txt_fecha_factura = Date
            Me.txt_importe = ""
            Me.txt_importe_aplicar = ""
            Me.txt_nombre_agente = ""
            Me.txt_nombre_agente_factura = ""
            Me.txt_nombre_cliente = ""
            Me.txt_numero = ""
            Me.txt_saldo = ""
            Me.txt_serie = ""
         End If
         rs.Close
      End If
      Me.frm_busqueda.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_busqueda.Visible = False
   End If
End Sub

Private Sub txt_busqueda_folio_LostFocus()
   var_ventana = 0
   Me.frm_busqueda.Visible = False
End Sub

Private Sub txt_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_deposito_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_aceptar_pedidos.SetFocus
   End If
End Sub

Private Sub txt_documento_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_documento_LostFocus()
   If Trim(txt_documento) <> "" Then
      If Trim(txt_documento) = "FA" Or Trim(txt_documento) = "NC" Or Trim(txt_documento) = "CH" Or Trim(txt_documento) = "CR" Then
      Else
         If Trim(txt_documento) = "SALDO" Then
            Me.txt_serie = ""
            Me.txt_numero = 0
            lv_lista.ListItems.Clear
            rs.Open "select distinct vcha_cli_clave_id, vcha_cli_nombre from tb_clientes where vcha_age_agente_id = '" + txt_agente + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockBatchOptimistic
            Dim list_item As ListItem
            Dim var_contador_lista As Integer
            If Not rs.EOF Then
               While Not rs.EOF
                     var_contador_lista = var_contador_lista + 1
                     Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
                     list_item.SubItems(1) = Trim(IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE))
                     rs.MoveNext
               Wend
            End If
            frm_lista.Visible = True
            lv_lista.SetFocus
            var_tipo_lista = 4
            rs.Close
         Else
            MsgBox "Documento incorrecto", vbOKOnly, "ATENCION"
            txt_documento = ""
         End If
      End If
   End If
End Sub

Private Sub txt_fecha_cheque_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha_cheque) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha_cheque)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fecha_cheque = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub



Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes.Value = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_importe_aplicar_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_agente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from TB_AGENTES order by vcha_AGE_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_AGE_AGENTE_ID)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_AGE_NOMBRE), "", rs!VCHA_AGE_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "AGENTES"
      var_tipo_lista = 3
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

Private Sub txt_nombre_agente_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_banco_cheque_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL CHEQUE"
      var_tipo_lista = 2
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

Private Sub txt_nombre_banco_cheque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_banco_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_bancos where vcha_ban_banco_id = '20' or vcha_ban_banco_id = '11' or vcha_ban_banco_id = '10'  or vcha_ban_banco_id = '22' order by vcha_ban_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_ban_banco_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_BAN_NOMBRE), "", rs!VCHA_BAN_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "BANCO DEL DEPOSITO"
      var_tipo_lista = 1
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

Private Sub txt_nombre_banco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_numero_deposito_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_numero_LostFocus()
   If Me.txt_documento = "SALDO" Then
   Else
      If Trim(txt_documento) <> "" Then
         If IsNumeric(txt_numero) Then
            rs.Open "SELECT * FROM TB_ENCABEZADO_cARTERA WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + txt_documento + "' AND VCHA_sER_SERIE_ID = '" + txt_serie + "' AND INTE_cAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               txt_importe = IIf(IsNull(rs!floa_Car_importe_neto), 0, rs!floa_Car_importe_neto) / IIf(IsNull(rs!floa_car_tipo_cambio), 1, rs!floa_car_tipo_cambio)
               txt_agente_factura = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
               txt_cliente = IIf(IsNull(rs!vcha_cli_clave_id), "", rs!vcha_cli_clave_id)
               If rsaux4.State = 1 Then
                  rsaux4.Close
               End If
               rsaux4.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + txt_agente_factura + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_agente_factura = IIf(IsNull(rsaux4!VCHA_AGE_NOMBRE), "", rsaux4!VCHA_AGE_NOMBRE)
               rsaux4.Close
               rsaux4.Open "SELECT * FROM TB_CLIENTES WHERE VCHA_CLI_CLAVE_ID = '" + txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
               txt_nombre_cliente = IIf(IsNull(rsaux4!VCHA_CLI_NOMBRE), "", rsaux4!VCHA_CLI_NOMBRE)
               rsaux4.Close
               rsaux4.Open "SELECT * FROM TB_SALDOS WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' AND VCHA_cAR_DOCUMENTO = '" + txt_documento + "' AND VCHA_sER_SERIE_ID = '" + txt_serie + "' AND INTE_cAR_NUMERO = " + txt_numero, cnn, adOpenDynamic, adLockOptimistic
               txt_saldo = IIf(IsNull(rsaux4!FLOA_sAL_IMPORTE), 0, rsaux4!FLOA_sAL_IMPORTE)
               rsaux4.Close
               txt_fecha_factura = IIf(IsNull(rs!dtim_Car_fecha), "", rs!dtim_Car_fecha)
            Else
               MsgBox "El documento no existe", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "Número de documento incorrecto", vbOKOnly, "ATENCION"
            txt_numero = ""
         End If
      Else
         MsgBox "Falta indicar el documento", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_relacion_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_relacion_LostFocus()
    If Trim(txt_relacion) <> "" Then
       rs.Open "select * from tb_relacion_cobranza with (nolock) where vcha_Rco_folio = '" + txt_relacion + "'", cnn, adOpenDynamic, adLockOptimistic
       Me.lv_relacion.ListItems.Clear
       If Not rs.EOF Then
          While Not rs.EOF
                Set list_item = lv_relacion.ListItems.Add(, , rs!vcha_Rco_folio)
                list_item.SubItems(1) = IIf(IsNull(rs!dtim_rco_fecha_relacion), "", rs!dtim_rco_fecha_relacion)
                list_item.SubItems(2) = IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID)
                rsaux.Open "SELECT * FROM TB_AGENTES WHERE VCHA_AGE_AGENTE_ID = '" + IIf(IsNull(rs!VCHA_AGE_AGENTE_ID), "", rs!VCHA_AGE_AGENTE_ID) + "'", cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   list_item.SubItems(3) = IIf(IsNull(rsaux!VCHA_AGE_NOMBRE), "", rsaux!VCHA_AGE_NOMBRE)
                Else
                   list_item.SubItems(3) = ""
                End If
                rsaux.Close
                list_item.SubItems(4) = IIf(IsNull(rs!VCHA_rCO_CHEQUE), "", rs!VCHA_rCO_CHEQUE)
                list_item.SubItems(5) = IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id)
                rsaux.Open "SELECT * FROM TB_BANCOS WHERE VCHA_BAN_BANCO_ID = '" + IIf(IsNull(rs!vcha_ban_banco_id), "", rs!vcha_ban_banco_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   list_item.SubItems(6) = IIf(IsNull(rsaux!VCHA_BAN_NOMBRE), "", rsaux!VCHA_BAN_NOMBRE)
                Else
                   list_item.SubItems(6) = ""
                End If
                rsaux.Close
                list_item.SubItems(7) = IIf(IsNull(rs!dtim_rco_Fecha_cheque), "", rs!dtim_rco_Fecha_cheque)
                list_item.SubItems(8) = IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento)
                list_item.SubItems(9) = IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID)
                list_item.SubItems(10) = IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)
                rsaux.Open "select * from tb_encabezado_cartera where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_Car_documento = '" + IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento) + "' and vcha_Ser_Serie_id = '" + IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID) + "' and inte_Car_numero  = " + CStr(IIf(IsNull(rs!inte_Car_numero), 0, rs!inte_Car_numero)), cnn, adOpenDynamic, adLockOptimistic
                If Not rsaux.EOF Then
                   var_agente = IIf(IsNull(rsaux!VCHA_AGE_AGENTE_ID), "", rsaux!VCHA_AGE_AGENTE_ID)
                   var_cliente = IIf(IsNull(rsaux!vcha_cli_clave_id), "", rsaux!vcha_cli_clave_id)
                   list_item.SubItems(11) = var_agente
                   rsaux2.Open "select * from tb_agentes where vcha_age_agente_id = '" + var_agente + "'", cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux2.EOF Then
                      list_item.SubItems(12) = IIf(IsNull(rsaux2!VCHA_AGE_NOMBRE), "", rsaux2!VCHA_AGE_NOMBRE)
                   Else
                      list_item.SubItems(12) = ""
                   End If
                   rsaux2.Close
                   rsaux2.Open "select * from tb_clientes where vcha_cli_clave_id = '" + var_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
                   list_item.SubItems(13) = var_cliente
                   If Not rsaux2.EOF Then
                      list_item.SubItems(14) = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
                   Else
                      list_item.SubItems(14) = IIf(IsNull(rsaux2!VCHA_CLI_NOMBRE), "", rsaux2!VCHA_CLI_NOMBRE)
                   End If
                   rsaux2.Close
                   list_item.SubItems(15) = IIf(IsNull(rsaux!dtim_Car_fecha), "", rsaux!dtim_Car_fecha)
                   list_item.SubItems(16) = IIf(IsNull(rsaux!floa_Car_importe_neto), "", rsaux!floa_Car_importe_neto)
                   rsaux2.Open "select * from tb_saldos where vcha_emp_empresa_id = '" + var_empresa + "' and vcha_car_documento = '" + IIf(IsNull(rs!vcha_Car_documento), "", rs!vcha_Car_documento) + "' and vcha_ser_serie_id = '" + IIf(IsNull(rs!VCHA_SER_SERIE_ID), "", rs!VCHA_SER_SERIE_ID) + "' and inte_Car_numero  = " + CStr(IIf(IsNull(rs!inte_Car_numero), "", rs!inte_Car_numero)), cnn, adOpenDynamic, adLockOptimistic
                   If Not rsaux2.EOF Then
                      list_item.SubItems(17) = IIf(IsNull(rsaux2!FLOA_sAL_IMPORTE), 0, rsaux2!FLOA_sAL_IMPORTE)
                   Else
                      list_item.SubItems(17) = 0
                   End If
                   rsaux2.Close
                   
                Else
                End If
                rsaux.Close
                list_item.SubItems(18) = Format(IIf(IsNull(rs!floa_rco_importe), 0, rs!floa_rco_importe), "###,###,##0.00")
                list_item.SubItems(19) = IIf(IsNull(rs!FLOA_RCO_DESCUENTO_OTORGADO), 0, rs!FLOA_RCO_DESCUENTO_OTORGADO)
                list_item.SubItems(20) = IIf(IsNull(rs!char_rco_aplicada), "", rs!char_rco_aplicada)
                rs.MoveNext
          Wend
          If lv_relacion.ListItems.Count > 0 Then
             lv_relacion.ListItems.Item(1).Selected = True
             Me.lv_relacion.SetFocus
          End If
       Else
          MsgBox "La relación de cobranza no existe", vbOKOnly, "ATENCION "
          Me.txt_agente = ""
          Me.txt_agente_factura = ""
          Me.txt_cliente = ""
          Me.txt_descuento = ""
          Me.txt_documento = ""
          Me.txt_fecha = Date
          Me.txt_fecha_factura = Date
          Me.txt_importe = ""
          Me.txt_importe_aplicar = ""
          Me.txt_nombre_agente = ""
          Me.txt_nombre_agente_factura = ""
          Me.txt_nombre_cliente = ""
          Me.txt_numero = ""
          Me.txt_saldo = ""
          Me.txt_serie = ""
       End If
       rs.Close
    End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_total_relacion_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
