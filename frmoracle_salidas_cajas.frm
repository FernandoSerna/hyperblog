VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_salidas_cajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salida de cajas"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_sellos 
      Height          =   2340
      Left            =   225
      TabIndex        =   0
      Top             =   960
      Width           =   3045
      Begin VB.Frame Frame4 
         Height          =   75
         Left            =   30
         TabIndex        =   5
         Top             =   645
         Width           =   2970
      End
      Begin VB.CommandButton cmd_cerrar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   690
         Picture         =   "frmoracle_salidas_cajas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cerrar Alt + C"
         Top             =   330
         Width           =   330
      End
      Begin VB.TextBox txt_sello 
         Height          =   315
         Left            =   585
         TabIndex        =   3
         Top             =   795
         Width           =   2385
      End
      Begin VB.CommandButton cmd_aceptar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         Picture         =   "frmoracle_salidas_cajas.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   330
         Width           =   330
      End
      Begin VB.CommandButton cmd_cancelar_sello 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   375
         Picture         =   "frmoracle_salidas_cajas.frx":024C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancelar Alt + C"
         Top             =   330
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_sellos 
         Height          =   1200
         Left            =   30
         TabIndex        =   6
         Top             =   1110
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   2117
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Número de Sello"
            Object.Width           =   5115
         EndProperty
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sello:"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Sellos"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   7
         Left            =   45
         TabIndex        =   7
         Top             =   135
         Width           =   2970
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_salidas_cajas.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   630
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6885
      Picture         =   "frmoracle_salidas_cajas.frx":0498
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Salir"
      Top             =   630
      Width           =   330
   End
   Begin VB.TextBox txt_clave_movimiento 
      Height          =   285
      Left            =   2205
      TabIndex        =   16
      Top             =   660
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   495
      Width           =   7200
   End
   Begin VB.Frame Frame2 
      Height          =   4185
      Left            =   105
      TabIndex        =   10
      Top             =   3150
      Width           =   7155
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1590
         TabIndex        =   11
         Top             =   435
         Width           =   3390
      End
      Begin MSComctlLib.ListView lv_salidas 
         Height          =   3090
         Left            =   15
         TabIndex        =   12
         Top             =   1035
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   5450
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "          Código"
            Object.Width           =   9172
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "O.S."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Número Caja"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Factura ceros"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Tipo_pedido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código de la Caja:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Lectura de Cajas"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   7080
      End
   End
   Begin VB.CommandButton cmd_cerrar_embarque 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmoracle_salidas_cajas.frx":0AD2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cerrar Embarque"
      Top             =   630
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   75
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":0BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":14AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":2324
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":2C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":34DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":3DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":3EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":3FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":40EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":41FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":430E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":4420
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":45C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":5414
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":55EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmoracle_salidas_cajas.frx":56FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Index           =   1
      Left            =   45
      TabIndex        =   19
      Top             =   870
      Width           =   7200
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   0
      Left            =   90
      TabIndex        =   26
      Top             =   945
      Width           =   2760
      Begin VB.TextBox txt_embarque 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   915
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   390
         Width           =   1620
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   105
         TabIndex        =   29
         Top             =   540
         Width           =   765
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Embarque"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   28
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   4
      Left            =   5040
      TabIndex        =   23
      Top             =   945
      Width           =   2205
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Cantidad Surtida"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   5
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Width           =   2130
      End
      Begin VB.Label lbl_recibidos 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         TabIndex        =   24
         Top             =   420
         Width           =   1860
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Index           =   3
      Left            =   2880
      TabIndex        =   20
      Top             =   945
      Width           =   2145
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Cantidad a Surtir"
         ForeColor       =   &H8000000E&
         Height          =   210
         Index           =   4
         Left            =   30
         TabIndex        =   22
         Top             =   120
         Width           =   2070
      End
      Begin VB.Label lbl_enviados 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   195
         TabIndex        =   21
         Top             =   420
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Index           =   1
      Left            =   105
      TabIndex        =   30
      Top             =   1800
      Width           =   7155
      Begin VB.TextBox txt_origen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   32
         Top             =   480
         Width           =   6150
      End
      Begin VB.TextBox txt_agente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   885
         TabIndex        =   31
         Top             =   825
         Width           =   6150
      End
      Begin VB.Label label 
         BackColor       =   &H000000C0&
         Caption         =   " Datos del Movimiento"
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   30
         TabIndex        =   35
         Top             =   120
         Width           =   7080
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Almacen:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   34
         Top             =   510
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agente:"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   33
         Top             =   855
         Width           =   555
      End
   End
   Begin VB.Label lblnombremovimiento 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   150
      TabIndex        =   36
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmoracle_salidas_cajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
    
Private Sub ilumina_grid()
    var_n = lv_salidas.ListItems.Count
    For var_i = 1 To var_n
        lv_salidas.ListItems.Item(var_i).Selected = True
        If Trim(lv_salidas.selectedItem.SubItems(6)) = "S" Then
           lv_salidas.ListItems.Item(var_i).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = False
           lv_salidas.ListItems.Item(var_i).ForeColor = &HFF&
           lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
           lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
           lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
           lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
           lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
           lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
        Else
           lv_salidas.ListItems.Item(var_i).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = False
           lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = False
           lv_salidas.ListItems.Item(var_i).ForeColor = &H80000008
           lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000008
           lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000008
           lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000008
           lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000008
           lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000008
           lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000008
        End If
    Next var_i
    If var_renglon > 0 Then
       If var_renglon <= var_n Then
          var_i = var_renglon
          lv_salidas.ListItems.Item(var_i).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(1).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(2).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(3).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(4).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(5).Bold = True
          lv_salidas.ListItems.Item(var_i).ListSubItems(6).Bold = True
          lv_salidas.ListItems.Item(var_i).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H8000&
          lv_salidas.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H8000&
       End If
    End If
    lv_salidas.Refresh
End Sub


Private Sub cmd_aceptar_sello_Click()
   If Trim(txt_sello) <> "" Then
      rs.Open "insert into tb_Sellos (inte_emb_embarque, vcha_Sel_Sello) values (" + Me.txt_embarque + ",'" + Me.txt_sello + "')", cnn, adOpenDynamic, adLockOptimistic
      Set list_item = lv_sellos.ListItems.Add(, , txt_sello)
      Me.txt_sello = ""
      Me.txt_sello.SetFocus
   Else
      MsgBox "No se indico un sello", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_cancelar_sello_Click()
   Me.frm_sellos.Visible = False
End Sub

Private Sub cmd_cerrar_Click()
Dim var_contador_errores As Integer
Dim clnt As New SoapClient30
Dim var_arreglo() As String
Dim var_container_id As String
Dim var_trip_id As String
Dim var_b As Boolean
Dim var_si_pedidos_cedis As Integer
var_si_pedidos_cedis = 0
      var_contador_errores = 0
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux1.State = 1 Then
      rsaux1.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   If rsaux7.State = 1 Then
      rsaux7.Close
   End If
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
   If rsaux9.State = 1 Then
      rsaux9.Close
   End If
   If rsaux10.State = 1 Then
      rsaux10.Close
   End If
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
   'GoTo icg:
   'GoTo VAR_COSTALES:
   strconsulta = "SELECT * FROM XXVIA_COMP_LEIDA_VS_PED_AFEC WHERE inte_emb_embarque = ? and CANTIDAD_LEIDA > CANTIDAD_PEDIDA_AFECTADA"
   With comandoORA
        .ActiveConnection = cnnoracle_4
        .CommandType = adCmdText
        .CommandText = strconsulta
        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
        .Parameters.Append parametro
   End With
   Set rsaux3 = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   var_posible_cerrado_comparacion = 1
   If Not rsaux3.EOF Then
      var_posible_cerrado_comparacion = 0
      cnn.BeginTrans
      rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM  TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
      Else
         var_consecutivo = 1
      End If
      rs.Close
      rsaux.Open "INSERT INTO TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      While Not rsaux3.EOF
            strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE inte_emb_embarque = ? and SOURCE_HEADER_NUMBER = ? AND SEGMENT1 = ? AND FLOA_SAL_cANTIDAD_LEIDA > 0"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux3!inte_emb_embarque))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux3!source_header_number))
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux3!SEGMENT1)
                 .Parameters.Append parametro
            End With
            Set rsaux2 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            VAR_CAJAS = ""
            While Not rsaux2.EOF
                  If VAR_CAJAS = "" Then
                     VAR_CAJAS = "CAJA: " + CStr(IIf(IsNull(rsaux2!INTE_PAQ_CAJA), 0, rsaux2!INTE_PAQ_CAJA)) + " CANTIDAD: " + CStr(IIf(IsNull(rsaux2!FLOA_SAL_CANTIDAD_LEIDA), 0, rsaux2!FLOA_SAL_CANTIDAD_LEIDA))
                  Else
                     VAR_CAJAS = VAR_CAJAS + ", CAJA: " + CStr(IIf(IsNull(rsaux2!INTE_PAQ_CAJA), 0, rsaux2!INTE_PAQ_CAJA)) + " CANTIDAD: " + CStr(IIf(IsNull(rsaux2!FLOA_SAL_CANTIDAD_LEIDA), 0, rsaux2!FLOA_SAL_CANTIDAD_LEIDA))
                  End If
                  rsaux2.MoveNext
            Wend
            rsaux2.Close
            var_cadena = "INSERT INTO TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS (INTE_TEM_CONSECUTIVO, EMBARQUE, PEDIDO, CODIGO, DESCRIPCION, CANTIDAD_PEDIDA, CANTIDAD_LEIDA, CAJAS )"
            var_cadena = var_cadena + " VALUES (" + CStr(var_consecutivo) + "," + CStr(rsaux3!inte_emb_embarque) + "," + CStr(rsaux3!source_header_number) + ",'" + rsaux3!SEGMENT1 + "', '" + rsaux3!item_description + "'," + CStr(rsaux3!CANTIDAD_PEDIDA_AFECTADA) + "," + CStr(rsaux3!cantidad_leida) + ",'" + VAR_CAJAS + "')"
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            rsaux3.MoveNext
      Wend
   End If
   rsaux3.Close
   If var_posible_cerrado_comparacion = 1 Then
      rs.Open "SELECT CHAR_EMB_ESTATUS FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
      VAR_ESTATUS = IIf(IsNull(rs(0).Value), "", rs(0).Value)
      rs.Close
      If VAR_ESTATUS = "E" Then
         var_si = MsgBox("¿Desea cerrar el embarque?", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_si = MsgBox("Confirmar el cerrado del embarque", vbYesNo, "ATENCION")
            If var_si = 6 Then
               x = 1
            Else
               x = 0
            End If
         Else
            x = 0
         End If
         If x = 1 Then
            rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
            'VAR_X_TRIP_ID = rs!arreglo_0
            'var_x_trip_name = rs!arreglo_1
            VAR_X_TRIP_ID = 0
            var_x_trip_name = "X"
            rs.Close
            If var_x_trip_name <> "" Then
               rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               If rs!tipo_embarque = 2 Then
                  rsaux.Open "select distinct source_header_number from xxvia_tb_salidas_CAJAS where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
               var_Cadena_pedidos = ""
               var_j = 0
               While Not rsaux.EOF
                     If var_Cadena_pedidos = "" Then
                        var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                     Else
                        var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                     End If
                     var_j = var_j + 1
                     rsaux.MoveNext
               Wend
               rsaux.Close
               'cambio blind no se puede
               'var_cadena_pedidos = "'480521','480527'"
               var_cadena = "SELECT HCAS.CUST_ACCOUNT_ID, HCAS.CUST_ACCT_SITE_ID as customer_id, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1 from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID  AND to_number(source_header_number) IN (" + var_Cadena_pedidos + ")"
               var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID  AND released_status = 'Y'"
               rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
               
               var_i = 0
               While Not rsaux.EOF
                     'rsaux3.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " AND source_header_number = " + CStr(CDbl(rsaux!SOURCE_HEADER_NUMBER)) + " AND DELIVERY_DETAIL_ID = " + CStr(rsaux!DELIVERY_DETAIL_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     strconsulta = "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = ? AND source_header_number = ? AND DELIVERY_DETAIL_ID = ?"
                     With comandoORA
                         .ActiveConnection = cnnoracle_4
                         .CommandType = adCmdText
                         .CommandText = strconsulta
                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                         .Parameters.Append parametro
                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux!source_header_number))
                         .Parameters.Append parametro
                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(rsaux!delivery_detail_id))
                         .Parameters.Append parametro
                     End With
                     Set rsaux3 = comandoORA.execute
                     Set comandoORA = Nothing
                     Set parametro = Nothing
                     If rsaux3.EOF Then
                        var_cadena = "INSERT INTO XXVIA_TB_SALIDAS_CAJAS (INTE_EMB_EMBARQUE, SOURCE_HEADER_NUMBER, SEGMENT1, FLOA_SAL_CANTIDAD_LEIDA, INVENTORY_ITEM_ID, DELIVERY_DETAIL_ID, SOURCE_LINE_NUMBER, DELIVERY_ID, INTE_PAQ_CAJA, CUSTOMER_ID, SUBINVENTORY, NAME, COLLECTOR_ID, ITEM_DESCRIPTION, CUSTOMER_NAME, char_paq_estatus)"
                        var_cadena = var_cadena + " values (" + Me.txt_embarque + "," + CStr(CDbl(rsaux!source_header_number)) + ",'" + rsaux!SEGMENT1 + "',0," + CStr(rsaux!inventory_item_id) + "," + CStr(rsaux!delivery_detail_id) + ",'" + CStr(rsaux!SOURCE_LINE_NUMBER) + "'," + CStr(IIf(IsNull(rsaux!delivery_id), 0, rsaux!delivery_id)) + ",0," + CStr(rsaux!CUSTOMER_ID) + ",'" + CStr(IIf(IsNull(rsaux!subinventory), "", rsaux!subinventory)) + "', '" + var_nombre_agente_str + "','" + CStr(VAR_AGENTE_str) + "','" + CStr(rsaux!Description) + "','" + rsaux!customer_name + "','S')"
                        rsaux4.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rsaux3.Close
                     rsaux.MoveNext
               Wend
               rsaux.Close
               If rsaux9.State = 1 Then
                 rsaux9.Close
               End If
               x = 1
               If x = 0 Then
                  rsaux9.Open "select * from tb_unidadesorganizacionales where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux9.EOF Then
                     VAR_USER_ID = rsaux9!user_id
                     VAR_RESP_ID = rsaux9!resp_id
                     VAR_RESP_APPL_ID = rsaux9!resp_appl_id
                  End If
                  rsaux9.Close
                  var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ")"
                  var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                  rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rsaux9.EOF
                        rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'MsgBox "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(rsaux9!header_id)) + ", " + CStr(CDbl(rsaux9!SOURCE_LINE_ID)) + ", 'PRODUCCION')"
                        On Error GoTo salir2:
                        rsaux7.Open "call XXVIA_SP_DEPURA_ORDEN_SURTIDO (" + CStr(CDbl(rsaux9!header_id)) + ", " + CStr(CDbl(rsaux9!source_LINE_ID)) + ", 'PRODUCCION'," + CStr(VAR_USER_ID) + "," + CStr(VAR_RESP_ID) + "," + CStr(VAR_RESP_APPL_ID) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux9.MoveNext
                  Wend
                  rsaux9.Close
                  rs.Close
               End If
      
                
      
      
      
               clnt.MSSoapInit var_webservice
               If rs.State = 1 Then
                  rs.Close
               End If
               x = 0
               If x = 1 Then
                  rs.Open "SELECT delivery_detail_id, sum(floa_sal_Cantidad_leida) as floa_sal_Cantidad_leida FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " group by delivery_detail_id ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  'rs.Open "SELECT delivery_detail_id, sum(floa_sal_Cantidad_leida) as floa_sal_Cantidad_leida FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque + " and source_header_number IN (480521,480527) group by delivery_detail_id ", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  While Not rs.EOF
                        'rsaux.Open "SELECT * FROM WSH_DELIVERABLES_V WHERE delivery_detail_id = " + CStr(rs!DELIVERY_DETAIL_ID) + " AND RELEASED_STATUS = 'Y'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'If Not rsaux.EOF Then
                        'var_b = clnt.actualizar_detalle(Val(rs!delivery_detail_id), CDbl(rs!FLOA_sAL_cANTIDAD_LEIDA), "OE", 0)
                        On Error GoTo salir2:
                        rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        rsaux6.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                        'rsaux6.Open "select max(inte_paq_caja) as inte_paq_caja  from xxvia_tb_Salidas_cajas where delivery_detail_id = " + CStr(rs!DELIVERY_DETAIL_ID), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        strconsulta = "select max(inte_paq_caja) as inte_paq_caja  from xxvia_tb_Salidas_cajas where delivery_detail_id = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!delivery_detail_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux6 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_consecutivo = rsaux6!INTE_PAQ_CAJA
                        rsaux6.Close
                        If rs!delivery_detail_id = 16566643 Then
                           var_consecutivo = var_consecutivo
                        End If
                        
                        
                        strconsulta = "SELECT source_header_type_name  FROM WSH_DELIVERABLES_V  WHERE delivery_detail_id = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 10, rs!delivery_detail_id)
                             .Parameters.Append parametro
                        End With
                        Set rsaux6 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        var_nombre_tipo_pedido = rsaux6!source_header_type_name
                        rsaux6.Close
                        If var_nombre_tipo_pedido = "VIA_MAYOREO_MTY" Then
                           var_si_pedidos_cedis = 1
                           rsaux6.Open "CALL xxvia_pk_interfaces_om.actualizar_detalle (1.0, " + CStr(rs!delivery_detail_id) + ",0,'OE'," + CStr(var_consecutivo) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_contador_errores = 0
                        Else
                           'MsgBox "CALL xxvia_pk_interfaces_om.actualizar_detalle (1.0, " + CStr(rs!delivery_detail_id) + "," + CStr(rs!FLOA_SAL_CANTIDAD_LEIDA) + ",'OE'," + CStr(var_consecutivo) + ")"
                           
                           strconsulta = "CALL xxvia_pk_interfaces_om.actualizar_detalle (1.0, ?,?,'OE',?)"
                           With comandoORA
                                .ActiveConnection = cnnoracle_4
                                .CommandType = adCmdText
                                .CommandText = strconsulta
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!delivery_detail_id)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rs!FLOA_SAL_CANTIDAD_LEIDA)
                                .Parameters.Append parametro
                                Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_consecutivo)
                                .Parameters.Append parametro
                           End With
                           Set rsaux6 = comandoORA.execute
                           Set comandoORA = Nothing
                           Set parametro = Nothing
                           'rsaux6.Open "CALL xxvia_pk_interfaces_om.actualizar_detalle (1.0, " + CStr(rs!delivery_detail_id) + "," + CStr(rs!FLOA_SAL_CANTIDAD_LEIDA) + ",'OE'," + CStr(var_consecutivo) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           var_contador_errores = 0
                        End If
                        'End If
                        'rsaux.Close
                        rs.MoveNext
                  Wend
                  rs.Close
               Else
                  On Error GoTo salir2:
                  
                  var_cadena = "select distinct source_header_number from xxvia_tb_salidas_cajas where  inte_emb_EMBARQUE = ? "
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = var_cadena
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                       .Parameters.Append parametro
                  End With
                  Set rsaux6 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                                          
                  While Not rsaux6.EOF
                        On Error GoTo salir2:
                        var_cadena = "call xxvia_sp_act_det_pedido_2 (?)"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        
                        var_cadena = "select order_type_id from oe_order_headers_all where order_number = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = var_cadena
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(rsaux6!source_header_number))
                             .Parameters.Append parametro
                        End With
                        Set rsaux7 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        If rsaux7!ORDER_TYPE_ID = 1002 Or rsaux7!ORDER_TYPE_ID = 1023 Then
                           rsaux7.Close
                           rsaux9.Open "SELECT * FROM TB_ORACLE_PEDIDOS_CERRADOS_CN WHERE PEDIDO = " + CStr(rsaux6!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                           If rsaux9.EOF Then
                              rsaux10.Open "INSERT INTO TB_ORACLE_PEDIDOS_CERRADOS_CN (PEDIDO) VALUES ('" + CStr(rsaux6!source_header_number) + "')", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux9.Close
                        
                        Else
                           rsaux7.Close
                           rsaux9.Open "SELECT * FROM TB_ORACLE_PEDIDOS_CERRADOS WHERE PEDIDO = " + CStr(rsaux6!source_header_number), cnn, adOpenDynamic, adLockOptimistic
                           If rsaux9.EOF Then
                              rsaux10.Open "INSERT INTO TB_ORACLE_PEDIDOS_CERRADOS (PEDIDO, REQUEST_ID) VALUES (" + CStr(rsaux6!source_header_number) + ",0)", cnn, adOpenDynamic, adLockOptimistic
                           End If
                           rsaux9.Close
                        End If
                        rsaux6.MoveNext
                  Wend
                  rsaux6.Close
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                
                  
                  
                  
                  
                  'rsaux6.Open "CALL xxvia_sp_act_det_pedido (" + Me.txt_embarque + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
               End If
               Set clnt = Nothing
               
               
               'clnt.MSSoapInit var_webservice
               'rs.Open "SELECT DISTINCT DELIVERY_ID FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               'While Not rs.EOF
               '
               '      var_arreglo = clnt.ASIGNAR_embarque(rs!delivery_id, Val(VAR_X_TRIP_ID), "CONFIRM")
               '      rs.MoveNext
               'Wend
               'rs.Close
               'Set clint = Nothing
               x = 1
               If x = 0 Then
               rs.Open "SELECT * FROM XXVIA_TB_SALIDAS_CAJAS WHERE INTE_EMB_EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               While Not rs.EOF
                     If IIf(IsNull(rs!FLOA_SAL_CANTIDAD_LEIDA), 0, rs!FLOA_SAL_CANTIDAD_LEIDA) > 0 Then
                        var_cadena = "INSERT INTO XXVIA_TB_DETALLE_CAJAS (EMBARQUE, PEDIDO,AGENTE, NOMBRE_AGENTE,CLIENTE,NOMBRE_CLIENTE,CODIGO, DESCRIPCION, CANTIDAD, PESO, CAJA, INVENTORY_ITEM_ID, CAJA_PEDIDO)"
                        var_cadena = var_cadena + " values (" + Me.txt_embarque + ", " + CStr(rs!source_header_number) + ",'" + CStr(IIf(IsNull(rs!collector_id), 0, rs!collector_id)) + "', '" + IIf(IsNull(rs!Name), "", rs!Name) + "',  '" + CStr(rs!CUSTOMER_ID) + "','" + IIf(IsNull(rs!customer_name), "", rs!customer_name) + "','" + rs!SEGMENT1 + "','" + rs!item_description + "'," + CStr(rs!FLOA_SAL_CANTIDAD_LEIDA) + ",0," + CStr(rs!INTE_PAQ_CAJA) + "," + CStr(rs!inventory_item_id) + "," + CStr(IIf(IsNull(rs!caja_pedido), 0, rs!caja_pedido)) + ")"
                        rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     End If
                     rs.MoveNext
               Wend
               rs.Close
               End If
               rs.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET CHAR_EMB_ESTATUS = 'I', FECHA_FIN = SYSDATE WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
               rs.Open "UPDATE TB_ORACLE_EMBARQUES_ORDENES SET estatus = 'I' WHERE inte_emb_EMBARQUE = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
               x = 0
               If x = 1 Then
                  rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
                        If rs!tipo_embarque = 2 Then
                           rsaux.Open "select distinct source_header_number from xxvia_tb_salidas_cAJAS where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        var_Cadena_pedidos = ""
                        var_j = 0
                        While Not rsaux.EOF
                              If var_Cadena_pedidos = "" Then
                                 var_Cadena_pedidos = "'" + CStr(rsaux!source_header_number) + "'"
                              Else
                                 var_Cadena_pedidos = var_Cadena_pedidos + ", '" + CStr(rsaux!source_header_number) + "'"
                              End If
                              var_j = var_j + 1
                              rsaux.MoveNext
                        Wend
                        rsaux.Close
                        var_i = 0
                        If var_i = 1 Then
                           While var_j <> var_i
                                 var_i = 0
                                 var_cadena = "SELECT e.collector_id, A.SOURCE_HEADER_NUMBER,  HL.ADDRESS1 AS CUSTOMER_NAME,  A.released_status,  E.NAME , sum(shipped_quantity) as cantidad from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, hz_customer_profiles D, ar_collectors e Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id "
                                 var_cadena = var_cadena + " AND A.SOURCE_HEADER_NUMBER in (" + var_Cadena_pedidos + ") AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND D.collector_id = e.collector_id AND released_status = 'C' group by  e.collector_id, A.SOURCE_HEADER_NUMBER, HL.ADDRESS1,  A.released_status,  E.NAME"
                                 'MsgBox var_cadena_pedidos
                                 rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 While Not rsaux.EOF
                                       var_i = var_i + 1
                                       rsaux.MoveNext
                                 Wend
                                 rsaux.Close
                           Wend
                           x = 1
                           If x = 0 Then
                              var_cadena_pedidos_global = var_Cadena_pedidos
                              var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ") "
                              var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                              rsaux7.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux7.EOF Then
                                 var_tipo_depurado = 1
                                 frmoracle_depurar_pedidos.Show 1
                              End If
                              rsaux7.Close
                              var_tipo_depurado = 0
                              var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                              var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                              rsaux9.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 rsaux9.Close
                                 var_sigue = 1
                                 While var_sigue = 1
                                       If rsaux8.State = 1 Then
                                          rsaux8.Close
                                       End If
                                       var_cadena = "SELECT a.source_line_id, OHA.HEADER_ID, HCAS.CUST_ACCT_SITE_ID, HCAS.PARTY_SITE_ID,HPS.LOCATION_ID, HL.ADDRESS1 AS CUSTOMER_NAME, A.inventory_item_id,a.date_requested,A.source_header_number,A.delivery_id,A.delivery_detail_id,A.organization_id,A.subinventory,A.delivery_line_id,A.inventory_item_id,C.DESCRIPTION,A.source_line_number,A.requested_quantity,A.released_status, c.segment1, E.NAME from hz_cust_acct_sites_all HCAS, HZ_PARTY_SITES HPS, HZ_LOCATIONS HL, OE_ORDER_HEADERS_ALL OHA, WSH_DELIVERABLES_V A, HZ_CUST_SITE_USES_ALL HCSU, xxvia_system_items_b C, hz_customer_profiles D, ar_collectors E Where HCAS.PARTY_SITE_ID = HPS.PARTY_SITE_ID AND HPS.LOCATION_ID =HL.LOCATION_ID AND HCSU.SITE_USE_ID= OHA.INVOICE_TO_ORG_ID AND HCSU.SITE_USE_ID = D.site_use_id AND A.SOURCE_HEADER_NUMBER in (" + var_cadena_pedidos_global + ")"
                                       var_cadena = var_cadena + " AND A.SOURCE_HEADER_ID = OHA.HEADER_ID AND HCSU.CUST_ACCT_SITE_ID = HCAS.CUST_ACCT_SITE_ID AND A.inventory_item_id  = c.inventory_item_id AND A.ORGANIZATION_ID = C.ORGANIZATION_ID AND D.collector_id = e.collector_id AND released_status = 'B' order by A.source_header_number"
                                       rsaux8.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                       If rsaux8.EOF Then
                                          var_sigue = 0
                                       Else
                                          While Not rsaux8.EOF
                                                rsaux7.Open "SELECT * FROM TB_ORACLE_NEGADO WHERE PEDIDO IN (" + CStr(rsaux8!source_header_number) + ") AND INVENTORY_ITEM_ID = " + CStr(rsaux8!inventory_item_id), cnn, adOpenDynamic, adLockOptimistic
                                                rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                                Set clnt = Nothing
                                                clnt.MSSoapInit var_webservice
                                                var_s = clnt.cancelar_back_order(CDbl(rsaux8!header_id), CDbl(rsaux8!source_LINE_ID), rsaux7!CAUSA_NEGADO)
                                                Set clnt = Nothing
                                                rsaux7.Close
                                                rsaux8.MoveNext
                                          Wend
                                       End If
                                       rsaux8.Close
                                 Wend
                              Else
                                 rsaux9.Close
                              End If
                           End If 'x
                        End If
                     End If
                  End If
               End If
               '--------------- confirmar pedidos
               x = 1
               If x = 1 Then
                  
                  rsaux.Open "alter session set nls_language = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If rs.State = 1 Then
                     rs.Close
                  End If
                  rs.Open "SELECT * FROM XXVIA_TB_ENCABEZADO_EMBARQUES WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     VAR_X_TRIP_ID = 1
                     var_x_trip_name = "X"
                     VAR_ESTATUS = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
                     If IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus) = "I" Then
                        If rs!tipo_embarque = 1 Then
                           rsaux.Open "select distinct source_header_number from xxvia_tb_salidas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        If rs!tipo_embarque = 2 Then
                           rsaux.Open "select distinct source_header_number from xxvia_tb_SAlidas_cajas where inte_emb_embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                        End If
                        VAR_CADENA_PEDIDOS_M = ""
                        While Not rsaux.EOF
                              If VAR_CADENA_PEDIDOS_M = "" Then
                                 VAR_CADENA_PEDIDOS_M = CStr(rsaux!source_header_number)
                              Else
                                 VAR_CADENA_PEDIDOS_M = VAR_CADENA_PEDIDOS_M + ", " + CStr(rsaux!source_header_number)
                              End If
                              rsaux.MoveNext
                        Wend
                        var_Cadena_pedidos = ""
                        rsaux.MoveFirst
                        While Not rsaux.EOF
                              If rsaux1.State = 1 Then
                                 rsaux1.Close
                              End If
                              rsaux1.Open "select distinct delivery_id from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              'rsaux1.Open "SELECT delivery_id FROM wsh_delivery_details wdd, wsh_delivery_assignments wda WHERE NVL(type, 'S')      IN ('S','C') AND wda.delivery_detail_id = wdd.delivery_detail_id and wdd.SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " and wdd.org_id = " + var_empresa + " and wdd.organization_id =" + var_unidad_organizacional + " group by delivery_id", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              VAR_ENTREGA = rsaux1!delivery_id
                              rsaux1.Close
                              rsaux1.Open "select distinct source_header_number from wsh_deliverables_v where delivery_id = " + CStr(VAR_ENTREGA), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              If Not rsaux1.EOF Then
                                 var_j = 0
                                 While Not rsaux1.EOF
                                       var_j = var_j + 1
                                       rsaux1.MoveNext
                                 Wend
                                 If var_j > 1 Then
                                    If var_Cadena_pedidos = "" Then
                                       var_Cadena_pedidos = CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                                    Else
                                       var_Cadena_pedidos = var_Cadena_pedidos + ", " + CStr(rsaux!source_header_number) + " ENTREGA: " + CStr(VAR_ENTREGA)
                                    End If
                                 End If
                              End If
                              rsaux1.Close
                              rsaux.MoveNext
                        Wend
                        rsaux.MoveFirst
                        If var_Cadena_pedidos <> "" Then
                           MsgBox "Los pedidos siguientes tienen dos entregas " + var_Cadena_pedidos
                        Else
                           cnn.BeginTrans
                           rsaux8.Open "SELECT MAX(CONSECUTIVO) FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux8.EOF Then
                              var_consecutivo = IIf(IsNull(rsaux8(0).Value), 0, rsaux8(0).Value) + 1
                           Else
                              var_consecutivo = 1
                           End If
                           rsaux8.Close
                           rsaux8.Open "insert into TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                           cnn.CommitTrans
                           rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!inte_emb_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, '" + CStr(rsaux2!FECHA_INICIO) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux2.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           rsaux8.Open "SELECT inte_emb_embarque, SOURCE_HEADER_NUMBER, SUM(FLOA_sal_cANTIDAD_LEIDA) AS CANTIDAD FROM XXVIA_TB_sALIDAS_CAJAS where source_header_number in (" + VAR_CADENA_PEDIDOS_M + ") GROUP BY inte_emb_embarque, SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 rsaux2.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + CStr(rsaux8!inte_emb_embarque), cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux2.EOF Then
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, '" + CStr(Date) + "'," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 Else
                                    rsaux10.Open "insert INTO TB_ORACLE_COMPARACION_PEDIDO_AFECTACION (PEDIDO, CANTIDAD_LEIDA, CANTIDAD_AFECTADA, fecha, CONSECUTIVO) VALUES (" + CStr(rsaux8!source_header_number) + "," + CStr(rsaux8!cantidad) + ",0, ''," + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux2.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           rsaux8.Open "SELECT pedido FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION WHERE CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                           While Not rsaux8.EOF
                                 rsaux1.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 rsaux10.Open "SELECT SOURCE_HEADER_NUMBER, SUM(SHIPPED_QUANTITY) AS CANTIDAD FROM WSH_DELIVERABLES_V WHERE SOURCE_HEADER_NUMBER = " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido)) + " GROUP BY SOURCE_HEADER_NUMBER", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                 If Not rsaux10.EOF Then
                                    rsaux1.Open "UPDATE TB_ORACLE_COMPARACION_PEDIDO_AFECTACION SET CANTIDAD_AFECTADA = " + CStr(IIf(IsNull(rsaux10!cantidad), 0, rsaux10!cantidad)) + " WHERE PEDIDO = " + CStr(rsaux8!pedido) + " AND CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                                 End If
                                 rsaux10.Close
                                 rsaux8.MoveNext
                           Wend
                           rsaux8.Close
                           rsaux8.Open "SELECT *  FROM TB_ORACLE_COMPARACION_PEDIDO_AFECTACION where cantidad_afectada > 0 and CANTIDAD_LEIDA <> cantidad_afectada AND CONSECUTIVO = " + CStr(var_consecutivo) + " order by PEDIDO desc "
                           If Not rsaux8.EOF Then
                              var_cadena_pedidos_mal = ""
                              While Not rsaux8.EOF
                                    If var_cadena_pedidos_mal = "" Then
                                       var_cadena_pedidos_mal = CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                                    Else
                                       var_cadena_pedidos_mal = var_cadena_pedidos_mal + ", " + CStr(IIf(IsNull(rsaux8!pedido), 0, rsaux8!pedido))
                                    End If
                                    rsaux8.MoveNext
                              Wend
                              MsgBox "Los siguientes pedidos tienen errores entra la cantidad leida y la cantidad afectada: " + CStr(var_cadena_pedidos_mal), vbOKOnly, "ATENCION"
                           Else
                              'clnt.MSSoapInit "http://intranet/WSOracle/wsInterfaceOM.asmx?wsdl"
                              If var_prueba = 0 Then
                                 clnt.MSSoapInit "http://intranet/WsEBS12Prod/wsInterfaceOM.asmx?wsdl"
                              Else
                                 If var_prueba = 2 Then
                                    clnt.MSSoapInit "http://intranet/WsEBS12Test/wsInterfaceOM.asmx?wsdl"
                                 End If
                              End If
                              
                              While Not rsaux.EOF
                                    rsaux2.Open "select distinct delivery_id, source_header_type_name from wsh_deliverables_v where SOURCE_HEADER_NUMBER = " + CStr(rsaux!source_header_number) + " AND delivery_id IS NOT NULL", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                    While Not rsaux2.EOF
                                          VAR_ENTREGA = rsaux2!delivery_id
                                          rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          VAR_ESTATUS = 0
                                          x = 1
                                          If x = 1 Then
                                             On Error GoTo salirc:
                                             var_nombre_tipo_pedido = rsaux2!source_header_type_name
                                             If var_nombre_tipo_pedido = "VIA_MAYOREO_MTY" Then
                                                rsaux6.Open "call XXVIA_SHIP_CONFIRM(" + CStr(VAR_ENTREGA) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                             Else
                                                rsaux6.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
'aqui se valida que no haya registros en blanco
                                                strconsulta = "select * from wsh_deliverables_v where delivery_id = ? and released_status = 'Y' and SHIPPED_QUANTITY is null"
                                                With comandoORA
                                                     .ActiveConnection = cnnoracle_4
                                                     .CommandType = adCmdText
                                                     .CommandText = strconsulta
                                                     Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, VAR_ENTREGA)
                                                     .Parameters.Append parametro
                                                 End With
                                                 Set rsaux6 = comandoORA.execute
                                                 Set comandoORA = Nothing
                                                 Set parametro = Nothing
                                                 
                                                 var_posible_entrega = 1
                                                 If Not rsaux6.EOF Then
                                                    var_posible_entrega = 0
                                                 End If
                                                 rsaux6.Close
                                                 If var_posible_entrega = 1 Then
                                                    strconsulta = "select sum(SHIPPED_QUANTITY) as cantidad from wsh_deliverables_v where source_header_number  = ?"
                                                    With comandoORA
                                                         .ActiveConnection = cnnoracle_4
                                                         .CommandType = adCmdText
                                                         .CommandText = strconsulta
                                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux!source_header_number)
                                                         .Parameters.Append parametro
                                                    End With
                                                    Set rsaux6 = comandoORA.execute
                                                    Set comandoORA = Nothing
                                                    Set parametro = Nothing
                                                    var_cantidad_oracle = IIf(IsNull(rsaux6!cantidad), 0, rsaux6!cantidad)
                                                    rsaux6.Close
                                                    
                                                    strconsulta = "select sum(floa_sal_cantidad_leida) as cantidad from xxvia_tb_Salidas_cajas where source_header_number  = ?"
                                                    With comandoORA
                                                         .ActiveConnection = cnnoracle_4
                                                         .CommandType = adCmdText
                                                         .CommandText = strconsulta
                                                         Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux!source_header_number)
                                                         .Parameters.Append parametro
                                                    End With
                                                    Set rsaux6 = comandoORA.execute
                                                    Set comandoORA = Nothing
                                                    Set parametro = Nothing
                                                    var_cantidad_sid = IIf(IsNull(rsaux6!cantidad), 0, rsaux6!cantidad)
                                                    rsaux6.Close
                                                    If var_cantidad_oracle <> var_cantidad_sid Then
                                                       var_posible_entrega = 0
                                                    End If
                                                                                                     
                                                 End If

'fin de la validacion
                                                If var_posible_entrega = 1 Then
                                                strconsulta = "CALL xxvia_pk_interfaces_om.asignar_embarque (?, ?, ?, ?, ?, ?)"
                                                With comandoORA
                             
                                                     .ActiveConnection = cnnoracle_4
                                  
                                                     .CommandType = adCmdText
                                                     .CommandText = strconsulta
                                                     Set parametro = .CreateParameter("p_api_version_number", adNumeric, adParamInput, 10, 1)
                                                     .Parameters.Append parametro
                                        
                                                     Set parametro = .CreateParameter("p_action_code", adVarChar, adParamInput, 10, "CONFIRM")
                                                     .Parameters.Append parametro
                                                     
                                                     Set parametro = .CreateParameter("p_delivery_id", adNumeric, adParamInput, 10, VAR_ENTREGA)
                                                     .Parameters.Append parametro
                                  
                                                     Set parametro = .CreateParameter("p_asg_trip_id", adNumeric, adParamInput, 200, CDbl(0))
                                                     .Parameters.Append parametro
                               
                                                     Set parametro = .CreateParameter("x_trip_id", adNumeric, adParamOutput, 10, Null)
                                                     .Parameters.Append parametro
                                  
                                                     Set parametro = .CreateParameter("x_trip_name", adVarChar, adParamOutput, 20000, Null)
                                                    .Parameters.Append parametro
                                                End With
                                                Set rsaux6 = comandoORA.execute
                                                
                                                Set comandoORA = Nothing
                                                Set parametro = Nothing
                                                Else
                                                   MsgBox "El pedido " + CStr(rsaux!source_header_number) + " no puede ser cerrado ya que hay diferencias en las cantidades", vbOKOnly, "ATENCION"
                                                   rsaux6.Open "UPDATE XXVIA_TB_ENCABEZADO_EMBARQUES SET CHAR_EMB_ESTATUS = 'E', FECHA_FIN = SYSDATE WHERE EMBARQUE = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
                                                End If
                                             End If
                                          Else
                                             On Error GoTo salirc:
                                             
                                             
                                             
                                             var_arreglo = clnt.ASIGNAR_embarque(VAR_ENTREGA, Val(VAR_X_TRIP_ID), "CONFIRM")
                                          End If
                                          rsaux1.Open "insert into tb_oracle_pedidos_confirmados (pedido, fecha, maquina, error) values (" + CStr(rsaux!source_header_number) + ", getdate(), '" + fun_NombrePc + "'," + CStr(VAR_ESTATUS) + ")", cnn, adOpenDynamic, adLockOptimistic
                                          rsaux2.MoveNext
                                    Wend
                                    rsaux2.Close
                                    rsaux.MoveNext
                              Wend
                              
                              Set clnt = Nothing
                              If var_posible_entrega = 1 Then
                                 MsgBox "Se termino de cerrar el embarque", vbOKOnly, "ATENCION"
                              Else
                                 MsgBox "EL EMBARQUE NO SE PUDO CERRAR", vbOKOnly, "ATENCION"
                              End If
                           End If
                           'HEINER
                           'Resume
                           If rsaux8.State = 1 Then
                              rsaux8.Close
                           End If
                        End If
                        If rsaux.State = 1 Then
                           rsaux.Close
                        End If
                     Else
                        If VAR_ESTATUS = "F" Then
                           MsgBox "EL embarque ya fue facturado"
                        Else
                           MsgBox "El embarque NO a sido cerrado", vbOKOnly, "ATENCION"
                        End If
                     End If
                  End If
                  If rs.State = 1 Then
                     rs.Close
                  End If
               End If
               '--------------- fin de confirmar pedidos
               
'inicio pedidos costales
               x = 1
               If x = 111111 Then
               If IsNumeric(Me.txt_embarque) Then
                  rs.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  strconsulta = "select * from xxvia_tb_encabezado_embarques where embarque = ?"
                  With comandoORA
                       .ActiveConnection = cnnoracle_4
                       .CommandType = adCmdText
                       .CommandText = strconsulta
                       Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                       .Parameters.Append parametro
                  End With
                  Set rsaux4 = comandoORA.execute
                  Set comandoORA = Nothing
                  Set parametro = Nothing
                  If Not rsaux4.EOF Then
                     VAR_eSTATUS_ = IIf(IsNull(rsaux4!char_emb_estatus), "", rsaux4!char_emb_estatus)
                     If VAR_eSTATUS_ <> "" Then
                        strconsulta = "select distinct source_header_number from XXVIA_TB_SALIDAS_cAJAS where inte_emb_embarque = ?"
                        With comandoORA
                             .ActiveConnection = cnnoracle_4
                             .CommandType = adCmdText
                             .CommandText = strconsulta
                             Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CDbl(Me.txt_embarque))
                             .Parameters.Append parametro
                        End With
                        Set rsaux5 = comandoORA.execute
                        Set comandoORA = Nothing
                        Set parametro = Nothing
                        While Not rsaux5.EOF
                              strconsulta = "SELECT TIPO_CAJA, COUNT(*) AS CANTIDAD FROM XXVIA_VW_CAJAS_POR_PEDIDO WHERE SOURCE_HEADER_NUMBER = ? AND (TIPO_CAJA LIKE '%COSTAL%' OR TIPO_CAJA LIKE 'CAJA BIASI') GROUP BY TIPO_CAJA"
                              With comandoORA
                                   .ActiveConnection = cnnoracle_4
                                   .CommandType = adCmdText
                                   .CommandText = strconsulta
                                   Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux5!source_header_number)
                                   .Parameters.Append parametro
                              End With
                              If rsaux6.State = 1 Then
                                 rsaux6.Close
                              End If
                              'MsgBox strconsulta
                              Set rsaux6 = comandoORA.execute
                              Set comandoORA = Nothing
                              Set parametro = Nothing
                  
                  
                              If Not rsaux6.EOF Then
                                 strconsulta = "select * from oe_order_headers_all where order_number = ?"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux5!source_header_number)
                                      .Parameters.Append parametro
                                 End With
                                 Set rsaux9 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 var_posible_pedido = 1
                                 If rsaux9!ORDER_TYPE_ID = 1002 Then
                                    var_posible_pedido = 0
                                    var_pedido_tienda = IIf(IsNull(rsaux9!order_number), "", rsaux9!order_number)
                                 End If
                                 rsaux9.Close
                                 If var_posible_pedido = 1 Then
                                    strconsulta = "SELECT * FROM OE_ORDER_HEADERS_ALL WHERE ORIG_SYS_DOCUMENT_REF = ?"
                                    With comandoORA
                                         .ActiveConnection = cnnoracle_4
                                         .CommandType = adCmdText
                                         .CommandText = strconsulta
                                         Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                                         .Parameters.Append parametro
                                    End With
                                    Set rsaux11 = comandoORA.execute
                                    Set comandoORA = Nothing
                                    Set parametro = Nothing
                                    If rsaux11.EOF Then
                                       strconsulta = "SELECT  distinct hp.party_name as nombre_titular,  account_number as vcha_tit_titular_id, hcas.cust_account_id AS VCHA_CLI_CLAVE_ID FROM hz_parties hp, hz_party_sites hps, hz_cust_accounts hca, hz_cust_acct_sites_all hcas, hz_cust_site_uses_all hcsu, hz_locations hl, hr_operating_units hr, hz_customer_profiles hcp, OE_ORDER_HEADERS_ALL OHA Where hca.party_id = hp.party_id AND hp.party_id = hps.party_id AND hps.party_site_id = hcas.party_site_id AND hca.cust_account_id = hcas.cust_account_id AND hca.cust_account_id = hcas.cust_account_id AND hcas.cust_acct_site_id = hcsu.cust_acct_site_id AND hps.location_id = hl.location_id AND hcas.org_id = hr.organization_id AND hcp.cust_account_id = hca.cust_account_id AND hcp.party_id = hp.party_id AND hcsu.site_use_id = hcp.site_use_id AND hcas.cust_account_id = OHA.SOLD_TO_ORG_ID AND ORDER_NUMBER = ? ORDER BY hp.party_name"
                                       With comandoORA
                                            .ActiveConnection = cnnoracle_4
                                            .CommandType = adCmdText
                                            .CommandText = strconsulta
                                            Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, Trim(CStr((rsaux5!source_header_number))))
                                            .Parameters.Append parametro
                                       End With
                                       Set rsaux12 = comandoORA.execute
                                       Set comandoORA = Nothing
                                       Set parametro = Nothing
                                       If Not rsaux12.EOF Then
                                          rsaux13.Open "SELECT * FROM TB_ORACLE_TITULARES_FACTURA_COSTALES WHERE TITULAR = '" + CStr(rsaux12!vcha_tit_titular_id) + "'", cnn, adOpenDynamic, adLockOptimistic
                                          If Not rsaux13.EOF Then
                                             var_posible_pedido = 1
                                          Else
                                             var_posible_pedido = 0
                                          End If
                                          rsaux13.Close
                                       Else
                                          var_posible_pedido = 0
                                       End If
                                       rsaux12.Close
                                       If var_posible_pedido = 1 Then
                                          strconsulta = "SELECT SOLD_TO_ORG_ID AS TITULAR, SHIP_TO_ORG_ID AS ESTABLECIMIENTO, INVOICE_TO_ORG_ID AS CLIENTE, PRICE_LIST_ID AS LISTA_PRECIOS FROM OE_ORDER_HEADERS_ALL WHERE ORDER_NUMBER = ?"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, CStr(rsaux5!source_header_number))
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux7 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                              
                                          strconsulta = "select name from qp_secu_list_headers_v where list_header_id = ?"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rsaux7!LISTA_PRECIOS)
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux8 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          var_lista_precios = rsaux8!Name
                                          rsaux8.Close
                                          var_clave_tipo_pedido = 1681
                                          strconsulta = "INSERT INTO oe_headers_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref, creation_date, created_by, last_update_date, last_updated_by, operation_code , sold_to_org_id        , SHIP_TO_ORG_id                   ,INVOICE_TO_ORG_ID     , Order_type_ID, PRICE_LIST, SHIP_FROM_ORG_ID, attribute7)"
                                          strconsulta = strconsulta + "  VALUES (1001,?,SYSDATE,-1,SYSDATE,-1,'INSERT', ?,?,?,?,?,?,?)"
                                          With comandoORA
                                               .ActiveConnection = cnnoracle_4
                                               .CommandType = adCmdText
                                               .CommandText = strconsulta
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!TITULAR)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!ESTABLECIMIENTO)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux7!Cliente)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_clave_tipo_pedido)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_lista_precios)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_unidad_organizacional)
                                               .Parameters.Append parametro
                                               Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "FACT. DE COSTALES")
                                               .Parameters.Append parametro
                                          End With
                                          Set rsaux8 = comandoORA.execute
                                          Set comandoORA = Nothing
                                          Set parametro = Nothing
                                          var_i = 0
                                          While Not rsaux6.EOF
                                                var_i = var_i + 1
                                                rs.Open "select * from tb_oracle_empaques where empaque = '" + rsaux6!tipo_caja + "'", cnn, adOpenDynamic, adLockOptimistic
                                                If Not rs.EOF Then
                                                   strconsulta = "select PRIMARY_UOM_CODE, INVENTORY_ITEM_ID from xxvia_system_items_b where SEGMENT1 = ? AND ORGANIZATION_ID = ?"
                                                   With comandoORA
                                                        .ActiveConnection = cnnoracle_4
                                                        .CommandType = adCmdText
                                                        .CommandText = strconsulta
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, rs!codigo)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_unidad_organizacional)
                                                        .Parameters.Append parametro
                                                   End With
                                                   Set rsaux8 = comandoORA.execute
                                                   Set comandoORA = Nothing
                                                   Set parametro = Nothing
                                                   var_inventory_item_id = rsaux8!inventory_item_id
                                                   VAR_MEDIDA = rsaux8!PRIMARY_UOM_CODE
                                                   rsaux8.Close
                                                   strconsulta = "INSERT INTO oe_lines_iface_all (ORDER_SOURCE_ID, orig_sys_document_ref,orig_sys_line_ref,inventory_item_id,ordered_quantity, operation_code, created_by, creation_date, last_updated_by, last_update_date, unit_selling_price, unit_list_price, calculate_price_flag, PRICING_QUANTITY, PRICING_QUANTITY_UOM, ATTRIBUTE1, subinventory, org_id, ship_from_org_id)"
                                                   strconsulta = strconsulta + " VALUES (1001,?,?,?, ?,'INSERT', -1,SYSDATE, -1,SYSDATE,0,0,'Y', ?, ?,'0','CDI_ALMPT',?,?)"
                                                   With comandoORA
                                                        .ActiveConnection = cnnoracle_4
                                                        .CommandType = adCmdText
                                                        .CommandText = strconsulta
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, "SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))))
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_i)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_inventory_item_id)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux6!cantidad)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, rsaux6!cantidad)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, VAR_MEDIDA)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_empresa)
                                                        .Parameters.Append parametro
                                                        Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_unidad_organizacional)
                                                        .Parameters.Append parametro
                                                   End With
                                                   Set rsaux8 = comandoORA.execute
                                                   Set comandoORA = Nothing
                                                   Set parametro = Nothing
                                                End If
                                                rs.Close
                                                rsaux6.MoveNext
                                          Wend
                                          On Error GoTo salir2
                                          rsaux8.Open "INSERT INTO oe_actions_iface_all (order_source_ID, orig_sys_document_ref, operation_code) VALUES (1001, 'SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))) + "','BOOK_ORDER')", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          rsaux8.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          rsaux8.Open "  ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          rsaux8.Open "CALL XXVIA_PK_INTERFACES_OM.importar_pedido('SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))) + "'," + var_empresa + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          
                                          rsaux8.Open "select * from where orig_sys_document_ref = 'SIDCOSTALES_" + Trim(CStr((rsaux5!source_header_number))) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                                          If Not rsaux8.EOF Then
                                             rsaux8.Close
                                             rsaux8.Open "INSERT INTO TB_ORACLE_PEDIDOS_CERRADOS (PEDIDO, REQUEST_ID) VALUES (" + CStr(rsaux8!order_number) + ",0)", cnn, adOpenDynamic, adLockOptimistic
                                          Else
                                             rsaux8.Close
                                          End If
                                          
                                          
                                       End If
                                    End If
                                    rsaux11.Close
                                 End If
                              End If
                              rsaux6.Close
                              rsaux5.MoveNext
                        Wend
                        rsaux5.Close
                        MsgBox "Se a terminado el proceso de insercion de costales"
                     End If
                  Else
                     MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
                  End If
                  rsaux4.Close
               End If
               End If
'fin pedidos costales
               
               
               'var_embarque_costales = CDbl(Me.txt_embarque)
               'frmoracle_crear_pedidos_costales.Show 1
               
               
               
               Me.frm_sellos.Visible = False
               Me.txt_codigo.Enabled = False
            Else
               MsgBox "No se pudo crear el embarque en oracle", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Nno se cerro el embarque", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El embarque ya habia sido cerrado", vbOKOnly, "ATENCION"
      End If
   Else
      rsaux.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " and codigo is null", cnn, adOpenDynamic, adLockOptimistic
      MsgBox "El embarque tiene diferencias entre las piezas pedidas y las leidas", vbOKOnly, "ATENCION"
      Set reporte = appl.OpenReport(App.Path + "\rep_oracle_diferencias_pedido_leido.rpt")
      var_cadena = "{TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      reporte.RecordSelectionFormula = var_cadena
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de diferencias pedido contra leido"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      
      rsaux.Open "DELETE FROM TB_TEMP_ORACLE_COMPARACION_CANTIDADES_LEIDAS_VS_PEDIDAS_AFECTADAS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   End If
   Exit Sub
salir2:
   MsgBox Err.Description
   If Err.Number = -2147217900 Then
      'MsgBox Err.Description
      rsaux12.Open "ALTER SESSION SET NLS_LANGUAGE = 'AMERICAN'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      rsaux12.Open "ALTER SESSION SET NLS_DATE_FORMAT='DD/MM/YYYY'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If var_contador_errores < 6 Then
         var_contador_errores = var_contador_errores + 1
         'MsgBox Err.Description
         Resume
      Else
         var_contador_errores = 0
         Resume Next
      End If
   End If
salirc:
   If Err.Number = -2147467259 Or Err.Number = -2147217900 Then
      'MsgBox Err.Description
      'MsgBox Err.Number
      Resume Next
      VAR_ESTATUS = 1
   End If
   If Err.Number = 3021 Then
      Resume Next
      'MsgBox Err.Description
   End If
   
   'MsgBox Err.Number
   
End Sub

Private Sub cmd_cerrar_embarque_Click()
   var_posible = True
   
   If var_bandera_asignacion = 0 Then
      rs.Open "select * from xxvia_tb_salidas_cajas where inte_emb_embarque = " + Me.txt_embarque + " and char_paq_estatus <> 'S' and inte_paq_caja > 0", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   Else
      For var_j = 1 To lv_salidas.ListItems.Count
          lv_salidas.ListItems.Item(var_j).Selected = True
          If Trim(lv_salidas.selectedItem.SubItems(6)) <> "S" Then
             var_posible = False
          End If
      Next var_j
   End If
   If var_posible = True Then
      Me.lv_sellos.ListItems.Clear
      rs.Open "select * from tb_Sellos where inte_emb_embarque = " + Me.txt_embarque, cnn, adOpenDynamic, adLockOptimistic
      var_numero_items = 0
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = lv_sellos.ListItems.Add(, , rs!vcha_sel_Sello)
               rs.MoveNext
               var_numero_items = var_numero_items + 1
         Wend
      End If
      If var_numero_items > 5 Then
         lv_sellos.ColumnHeaders(1).Width = 2650
      Else
         lv_sellos.ColumnHeaders(1).Width = 2850
      End If
      rs.Close
      Me.frm_sellos.Visible = True
      Me.cmd_cerrar.SetFocus
   Else
      MsgBox "Faltan cajas por leer", vbOKOnly, "ATENCION"
   End If
End Sub


Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_posible_kanban = 0
   Top = 0
   Left = 1900
   var_cantidad_enviada = 0
   var_cantidad_recibida = 0
   Me.frm_sellos.Visible = False
   Me.txt_embarque = frmnumero_embarque.txt_embarque
   Dim var_posible As Boolean
   var_posible = True
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select subinventory, name, inte_emb_embarque, inte_paq_caja, source_header_number, char_paq_estatus, sum(floa_sal_cantidad_leida) as cantidad from XXVIA_TB_SALIDAS_CAJAS where inte_emb_embarque = " + Me.txt_embarque + "  group by subinventory,name, inte_emb_embarque, inte_paq_caja, source_header_number, char_paq_estatus", cnnoracle_4, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      Me.txt_agente = IIf(IsNull(rs!Name), "", rs!Name)
      Me.txt_origen = IIf(IsNull(rs!subinventory), "", rs!subinventory)
      While Not rs.EOF
            If var_bandera_asignacion <> 0 Then
               var_numero_caja = IIf(IsNull(INTE_PAQ_CAJA), 0, rs!INTE_PAQ_CAJA)
               If Len(Trim(Str(var_numero_caja))) = 1 Then
                  var_referencia_caja = "00" + Trim(Str(var_numero_caja))
               End If
               If Len(Trim(Str(var_numero_caja))) = 2 Then
                  var_referencia_caja = "0" + Trim(Str(var_numero_caja))
               End If
               If Len(Trim(Str(var_numero_caja))) = 3 Then
                  var_referencia_caja = Trim(Str(var_numero_caja))
               End If
               If Len(Trim(Str(txt_embarque))) = 1 Then
                  var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
               End If
               If Len(Trim(Str(txt_embarque))) = 2 Then
                  var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
               End If
               If Len(Trim(Str(txt_embarque))) = 3 Then
                  var_referencia_embarque = "000" + Trim(Str(txt_embarque))
               End If
               If Len(Trim(Str(txt_embarque))) = 4 Then
                   var_referencia_embarque = "00" + Trim(Str(txt_embarque))
               End If
               If Len(Trim(Str(txt_embarque))) = 5 Then
                  var_referencia_embarque = "0" + Trim(Str(txt_embarque))
               End If
               If Len(Trim(Str(txt_embarque))) = 6 Then
                  var_referencia_embarque = "" + Trim(Str(txt_embarque))
               End If
               var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
               Set list_item = lv_salidas.ListItems.Add(, , var_codigo_caja)
               list_item.SubItems(1) = IIf(IsNull(rs!cantidad), 0, rs!cantidad)
               list_item.SubItems(2) = IIf(IsNull(rs!source_header_number), 0, rs!source_header_number)
               list_item.SubItems(3) = IIf(IsNull(rs!INTE_PAQ_CAJA), 0, rs!INTE_PAQ_CAJA)
               list_item.SubItems(4) = 0
               list_item.SubItems(5) = ""
               list_item.SubItems(6) = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
               var_estatus_caja = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
               var_cantidad_enviada = var_cantidad_enviada + rs!cantidad
               If var_estatus_caja = "S" Then
                  var_cantidad_recibida = var_cantidad_recibida + rs!cantidad
               End If
               var_i = var_i + 1
            Else
               var_estatus_caja = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
               If var_estatus_caja = "S" Then
                  var_numero_caja = IIf(IsNull(INTE_PAQ_CAJA), 0, rs!INTE_PAQ_CAJA)
                  If Len(Trim(Str(var_numero_caja))) = 1 Then
                     var_referencia_caja = "00" + Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(var_numero_caja))) = 2 Then
                     var_referencia_caja = "0" + Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(var_numero_caja))) = 3 Then
                     var_referencia_caja = Trim(Str(var_numero_caja))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 1 Then
                     var_referencia_embarque = "00000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 2 Then
                     var_referencia_embarque = "0000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 3 Then
                     var_referencia_embarque = "000" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 4 Then
                      var_referencia_embarque = "00" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 5 Then
                     var_referencia_embarque = "0" + Trim(Str(txt_embarque))
                  End If
                  If Len(Trim(Str(txt_embarque))) = 6 Then
                     var_referencia_embarque = "" + Trim(Str(txt_embarque))
                  End If
                  var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
                  Set list_item = lv_salidas.ListItems.Add(, , var_codigo_caja)
                  list_item.SubItems(1) = IIf(IsNull(rs!cantidad), 0, rs!cantidad)
                  list_item.SubItems(2) = IIf(IsNull(rs!source_header_number), 0, rs!source_header_number)
                  list_item.SubItems(3) = IIf(IsNull(rs!INTE_PAQ_CAJA), 0, rs!INTE_PAQ_CAJA)
                  list_item.SubItems(4) = 0
                  list_item.SubItems(5) = ""
                  list_item.SubItems(6) = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
                  var_estatus_caja = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
                  var_cantidad_enviada = var_cantidad_enviada + rs!cantidad
                  If var_estatus_caja = "S" Then
                     var_cantidad_recibida = var_cantidad_recibida + rs!cantidad
                  End If
                  var_i = var_i + 1
                               
               End If
            End If
            rs.MoveNext
      Wend
      For var_j = 1 To Me.lv_salidas.ListItems.Count
          Me.lv_salidas.ListItems.Item(var_j).Selected = True
          If Me.lv_salidas.selectedItem.SubItems(6) = "S" Then
             lv_salidas.ListItems.Item(var_j).Bold = False
             lv_salidas.ListItems.Item(var_j).ListSubItems(1).Bold = False
             lv_salidas.ListItems.Item(var_j).ListSubItems(2).Bold = False
             lv_salidas.ListItems.Item(var_j).ListSubItems(3).Bold = False
             lv_salidas.ListItems.Item(var_j).ListSubItems(4).Bold = False
             lv_salidas.ListItems.Item(var_j).ListSubItems(5).Bold = False
             lv_salidas.ListItems.Item(var_j).ListSubItems(6).Bold = False
             lv_salidas.ListItems.Item(var_j).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HFF&
             lv_salidas.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HFF&
          End If
      Next var_j
      'For var_j = 1 To lv_salidas.ListItems.Count
      lbl_recibidos = Format(var_cantidad_recibida, "###,###,##0.00")
      lbl_enviados = Format(var_cantidad_enviada, "###,###,##0.00")
      
   End If
   rs.Close
   
   
End Sub

Private Sub lv_salidas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_salidas, ColumnHeader)
End Sub

Private Sub lv_sellos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
      rs.Open "delete from tb_sellos where vcha_emp_empresa_id = '" + var_empresa + "' and inte_emb_embarque = " + CStr(var_numero_embarque) + " and vcha_sel_sello = '" + lv_sellos.selectedItem + "'", cnn, adOpenDynamic, adLockOptimistic
      lv_sellos.ListItems.Remove (lv_sellos.selectedItem.Index)
   End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Trim(txt_codigo) <> "" Then
         If var_bandera_asignacion <> 0 Then
            Set itmfound = lv_salidas.findItem(Trim(txt_codigo), lvwText, , lvwPartial)
            If itmfound Is Nothing Then
               txt_codigo = ""
               frmmensaje.lbl_mensaje = "La caja no se encuentra en el embarque"
               frmmensaje.Show 1
               txt_codigo.SetFocus
               var_orden_surtido = 0
               var_caja = 0
               var_factura_ceros = 0
               var_tipo_pedido = ""
            Else
               itmfound.EnsureVisible
               itmfound.Selected = True
               If lv_salidas.selectedItem.SubItems(6) = "S" Then
                  frmmensaje.lbl_mensaje = "La caja ya fue surtida"
                  frmmensaje.Show 1
                  txt_codigo.SetFocus
                  var_orden_surtido = 0
                  var_caja = 0
                  var_factura_ceros = 0
                  var_tipo_pedido = ""
                  Me.txt_codigo = ""
                  Me.txt_codigo.SetFocus
               Else
                  var_orden_surtido = lv_salidas.selectedItem.SubItems(2)
                  var_caja = lv_salidas.selectedItem.SubItems(3)
                  var_factura_ceros = lv_salidas.selectedItem.SubItems(4)
                  var_tipo_pedido = lv_salidas.selectedItem.SubItems(5)
                  lv_salidas.selectedItem.SubItems(6) = "S"
                  var_embarque_auditar = CDbl(Me.txt_embarque)
                  var_caja_auditar = CDbl(var_caja)
                  If var_bandera_asignacion = 0 Then
                     frmoracle_audita_caja.Show 1
                     rs.Open "SELECT * FROM XXVIA_TB_cAJAS_AUDITADAS WHERE EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar) + " AND CANTIDAD_ORIGINAL <> CANTIDAD_AUDITADA", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If rs.EOF Then
                        rsaux.Open "update XXVIA_TB_SALIDAS_CAJAS set char_paq_estatus = 'S' where inte_emb_embarque = " + Me.txt_embarque + " and inte_paq_caja = " + CStr(var_caja), cnnoracle_4, adOpenDynamic, adLockOptimistic
                        var_renglon = lv_salidas.selectedItem.Index
                        Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) + CDbl(lv_salidas.selectedItem.SubItems(1)), "###,###,##0.00")
                        Call ilumina_grid
                        Me.txt_codigo = ""
                        Me.txt_codigo.SetFocus
                     Else
                        frmmensaje.lbl_mensaje = "Existen diferencias en la caja auditada"
                        frmmensaje.Show 1
                        txt_codigo.SetFocus
                        var_orden_surtido = 0
                        var_caja = 0
                        var_factura_ceros = 0
                        var_tipo_pedido = ""
                        Me.txt_codigo = ""
                        Me.txt_codigo.SetFocus
                     End If
                     rs.Close
                  Else
                     rsaux.Open "update XXVIA_TB_SALIDAS_CAJAS set char_paq_estatus = 'S' where inte_emb_embarque = " + Me.txt_embarque + " and inte_paq_caja = " + CStr(var_caja), cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_renglon = lv_salidas.selectedItem.Index
                     Me.lbl_recibidos = Format(CDbl(Me.lbl_recibidos) + CDbl(lv_salidas.selectedItem.SubItems(1)), "###,###,##0.00")
                     Call ilumina_grid
                     Me.txt_codigo = ""
                     Me.txt_codigo.SetFocus
                  End If
               End If
            End If
         Else
            var_referencia_embarque = Mid(Me.txt_codigo, 2, 6)
            var_referencia_caja = Mid(Me.txt_codigo, 8, 3)
            If IsNumeric(var_referencia_embarque) Then
               If IsNumeric(var_referencia_caja) Then
                  If CDbl(Me.txt_embarque) = CDbl(var_referencia_embarque) Then
                     rs.Open "select * from xxvia_tb_salidas_cajas where inte_emb_embarque  = " + CStr(var_referencia_embarque) + " and inte_paq_caja = " + CStr(var_referencia_caja) + " and floa_sal_cantidad_leida >0", cnnoracle_4, adOpenDynamic, adLockOptimistic
                     If Not rs.EOF Then
                        var_estatus_caja = IIf(IsNull(rs!char_paq_estatus), "", rs!char_paq_estatus)
                        If var_estatus_caja = "I" Then
                           var_embarque_auditar = CDbl(Me.txt_embarque)
                           var_caja_auditar = CDbl(var_referencia_caja)
                           frmoracle_audita_caja.Show 1
                           rsaux.Open "SELECT * FROM XXVIA_TB_cAJAS_AUDITADAS WHERE EMBARQUE = " + CStr(var_embarque_auditar) + " AND CAJA = " + CStr(var_caja_auditar) + " AND CANTIDAD_ORIGINAL <> CANTIDAD_AUDITADA", cnnoracle_4, adOpenDynamic, adLockOptimistic
                           If rsaux.EOF Then
                              If rsaux1.State = 1 Then
                                 rsaux1.Close
                              End If
                              frmoracle_sello_caja.Show 1
                              rsaux1.Open "update xxvia_tb_salidas_cajas set char_paq_estatus = 'S', sello = '" + var_sello_caja + "', AUDITADA = 1 where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              rsaux1.Open "select source_header_number, inte_paq_caja, char_paq_estatus, sum(floa_Sal_Cantidad_leida) as cantidad from xxvia_tb_salidas_cajas where inte_emb_embarque = " + CStr(Me.txt_embarque) + " and inte_paq_caja = " + CStr(var_caja_auditar) + " group by source_header_number, inte_paq_caja, char_paq_estatus", cnnoracle_4, adOpenDynamic, adLockOptimistic
                              var_codigo_caja = "C" + var_referencia_embarque + var_referencia_caja
                              Set list_item = lv_salidas.ListItems.Add(, , var_codigo_caja)
                              list_item.SubItems(1) = IIf(IsNull(rsaux1!cantidad), 0, rsaux1!cantidad)
                              list_item.SubItems(2) = IIf(IsNull(rsaux1!source_header_number), 0, rsaux1!source_header_number)
                              list_item.SubItems(3) = IIf(IsNull(rsaux1!INTE_PAQ_CAJA), 0, rsaux1!INTE_PAQ_CAJA)
                              list_item.SubItems(4) = 0
                              list_item.SubItems(5) = ""
                              list_item.SubItems(6) = IIf(IsNull(rsaux1!char_paq_estatus), "", rsaux1!char_paq_estatus)
                              var_estatus_caja = IIf(IsNull(rsaux1!char_paq_estatus), "", rsaux1!char_paq_estatus)
                              var_cantidad_enviada = var_cantidad_enviada + rsaux1!cantidad
                              If var_estatus_caja = "S" Then
                                 var_cantidad_recibida = var_cantidad_recibida + rsaux1!cantidad
                              End If
                              var_cantidad_leida = 0
                              For var_j = 1 To lv_salidas.ListItems.Count
                                  lv_salidas.ListItems.Item(var_j).Selected = True
                                  lv_salidas.ListItems.Item(var_j).Bold = False
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(1).Bold = False
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(2).Bold = False
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(3).Bold = False
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(4).Bold = False
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(5).Bold = False
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(6).Bold = False
                                  lv_salidas.ListItems.Item(var_j).ForeColor = &HFF&
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(1).ForeColor = &HFF&
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(2).ForeColor = &HFF&
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(3).ForeColor = &HFF&
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(4).ForeColor = &HFF&
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(5).ForeColor = &HFF&
                                  lv_salidas.ListItems.Item(var_j).ListSubItems(6).ForeColor = &HFF&
                                  var_cantidad_leida = var_cantidad_leida + CDbl(Me.lv_salidas.selectedItem.SubItems(1))
                                  var_i = var_i + 1
                              Next var_j
                              Me.lbl_recibidos = Format(var_cantidad_leida, "###,###,##0.00")
                              Me.txt_codigo = ""
                           Else
                              frmmensaje.lbl_mensaje = "Existen diferencias en la caja auditada"
                              frmmensaje.Show 1
                              rsaux8.Open "UPDATE XXVIA_TB_SALIDAS_CAJAS SET FLOA_SAL_cANTIDAD_LEIDA = 0, AUDITADA = 1, CHAR_PAQ_ESTATUS = '', SELLO = '' where inte_emb_embarque  = " + CStr(var_referencia_embarque) + " and inte_paq_caja = " + CStr(var_referencia_caja), cnnoracle_4, adOpenDynamic, adLockOptimistic
                              txt_codigo.SetFocus
                              var_orden_surtido = 0
                              var_caja = 0
                              var_factura_ceros = 0
                              var_tipo_pedido = ""
                              Me.txt_codigo = ""
                              Me.txt_codigo.SetFocus
                           End If
                           rsaux.Close
                        Else
                           If var_estatus_caja = "E" Then
                              frmmensaje.lbl_mensaje = "La caja no contiene información"
                              frmmensaje.Show 1
                              txt_codigo.SetFocus
                              var_orden_surtido = 0
                              var_caja = 0
                              var_factura_ceros = 0
                              var_tipo_pedido = ""
                              Me.txt_codigo = ""
                              Me.txt_codigo.SetFocus
                           End If
                           If var_estatus_caja = "S" Then
                              frmmensaje.lbl_mensaje = "La caja ya fue leida"
                              frmmensaje.Show 1
                              txt_codigo.SetFocus
                              var_orden_surtido = 0
                              var_caja = 0
                              var_factura_ceros = 0
                              var_tipo_pedido = ""
                              Me.txt_codigo = ""
                              Me.txt_codigo.SetFocus
                           End If
                           If var_estatus_caja = "" Then
                              frmmensaje.lbl_mensaje = "La caja no a sido cerrada aun"
                              frmmensaje.Show 1
                              txt_codigo.SetFocus
                              var_orden_surtido = 0
                              var_caja = 0
                              var_factura_ceros = 0
                              var_tipo_pedido = ""
                              Me.txt_codigo = ""
                              Me.txt_codigo.SetFocus
                           End If
                        End If
                     Else
                        frmmensaje.lbl_mensaje = "La caja no contiene información"
                        frmmensaje.Show 1
                        txt_codigo.SetFocus
                        var_orden_surtido = 0
                        var_caja = 0
                        var_factura_ceros = 0
                        var_tipo_pedido = ""
                        Me.txt_codigo = ""
                        Me.txt_codigo.SetFocus
                     End If
                     rs.Close
                  Else
                     frmmensaje.lbl_mensaje = "Número de embarque incorrecto"
                     frmmensaje.Show 1
                     txt_codigo.SetFocus
                     var_orden_surtido = 0
                     var_caja = 0
                     var_factura_ceros = 0
                     var_tipo_pedido = ""
                     Me.txt_codigo = ""
                     Me.txt_codigo.SetFocus
                  End If
               Else
                  frmmensaje.lbl_mensaje = "Número de caja incorrecto"
                  frmmensaje.Show 1
                  txt_codigo.SetFocus
                  var_orden_surtido = 0
                  var_caja = 0
                  var_factura_ceros = 0
                  var_tipo_pedido = ""
                  Me.txt_codigo = ""
                  Me.txt_codigo.SetFocus
               End If
            Else
               frmmensaje.lbl_mensaje = "Número de embarque incorrecto"
               frmmensaje.Show 1
               txt_codigo.SetFocus
               var_orden_surtido = 0
               var_caja = 0
               var_factura_ceros = 0
               var_tipo_pedido = ""
               Me.txt_codigo = ""
               Me.txt_codigo.SetFocus
            End If
         End If
      End If
   End If

End Sub

Private Sub txt_sello_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_aceptar_sello.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.frm_sellos.Visible = False
   End If
End Sub
