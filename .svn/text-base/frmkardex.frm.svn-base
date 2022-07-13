VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmkardex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frmkardex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6045
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmkardex.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   390
      Picture         =   "frmkardex.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5595
      Picture         =   "frmkardex.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   60
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Caption         =   " Busqueda de Artículo "
      Height          =   660
      Left            =   45
      TabIndex        =   20
      Top             =   1215
      Width           =   5880
      Begin VB.TextBox txt_descripcion 
         Height          =   300
         Left            =   1410
         TabIndex        =   6
         Top             =   270
         Width           =   4320
      End
      Begin VB.TextBox txt_codigo 
         Height          =   300
         Left            =   165
         TabIndex        =   5
         Top             =   270
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView mon_mes2 
      Height          =   2370
      Left            =   3030
      TabIndex        =   19
      Top             =   3915
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   73007105
      CurrentDate     =   37761
   End
   Begin MSComCtl2.MonthView mon_mes1 
      Height          =   2370
      Left            =   240
      TabIndex        =   18
      Top             =   3900
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   73007105
      CurrentDate     =   37761
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   120
      TabIndex        =   11
      Top             =   5910
      Width           =   5850
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   293
         Width           =   1305
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   293
         Width           =   1320
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   330
         Left            =   2385
         TabIndex        =   16
         Top             =   285
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Inicio del Perido"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   330
         Left            =   5175
         TabIndex        =   17
         Top             =   285
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Termino del Periodo"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3465
         TabIndex        =   15
         Top             =   353
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   525
         TabIndex        =   14
         Top             =   353
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Almacen "
      Height          =   660
      Left            =   60
      TabIndex        =   9
      Top             =   540
      Width           =   5880
      Begin VB.ComboBox cmb_almacen 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   225
         Width           =   5625
      End
      Begin VB.TextBox txt_almacen 
         Height          =   300
         Left            =   210
         TabIndex        =   10
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Artículos "
      Height          =   3990
      Left            =   90
      TabIndex        =   8
      Top             =   1890
      Width           =   5865
      Begin VB.CommandButton cmd_seleccion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1395
         Picture         =   "frmkardex.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_marcar 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   735
         Picture         =   "frmkardex.frx":131E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Marcar (Enter)"
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_invertir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Picture         =   "frmkardex.frx":1568
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_ninguno 
         Height          =   315
         Left            =   75
         Picture         =   "frmkardex.frx":163A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmd_todos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   405
         Picture         =   "frmkardex.frx":173C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_articulos 
         Height          =   3735
         Left            =   45
         TabIndex        =   7
         Top             =   180
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   6588
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
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7373
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Marca"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   6030
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1110
      Top             =   -15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":1952
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":222C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":2B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":30A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":397E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":4258
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":4B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":4C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":4D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":4E68
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   -15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":4F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":50FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":5322
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":557C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":56FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2235
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmkardex.frx":5924
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmkardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim var_maximo_kardex As Integer
Dim var_primera_vez As Boolean
Dim vx As String
Dim vy As String

Private Sub cmb_almacen_Click()
      txt_almacen = Obtener_llave(cnn, rs, "TB_almacenes", "VCHA_alm_NOMBRE", cmb_almacen, 2, "T")
End Sub

Private Sub cmb_almacen_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub cmd_imprimir_Click()
   Set TB_TEM_KARDEX_INV_INI = New TB_TEM_KARDEX_INV_INI
         If txt_almacen = "" Then
            MsgBox "No se a seleccionado ningun almacen", vbOKOnly, "ATENCION"
         Else
            cnn.CommandTimeout = 36000
            ok = TB_TEM_KARDEX_INV_INI.Anadir(var_maximo_kardex, txt_almacen, vx, vy, txt_inicio, txt_fin)
            rs.Open "SELECT * FROM VW_KARDEX WHERE VCHA_AUD_USUARIO = '" + vx + "' AND VCHA_AUD_MAQUINA = '" + vy + "' and int_kar_identificador = " + Str(var_maximo_kardex), cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               Set reporte = appl.OpenReport(App.Path + "\rep_kardex.rpt")
               reporte.RecordSelectionFormula = "{VW_KARDEX.VCHA_AUD_USUARIO} = '" + vx + "' and {VW_KARDEX.VCHA_AUD_MAQUINA} = '" + vy + "' and {vw_kardex.int_kar_identificador} = " + Str(var_maximo_kardex)
               frmvistasprevias.cr.ReportSource = reporte
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               frmvistasprevias.cr.ViewReport
               frmvistasprevias.Caption = "Reporte de movimientos Kardex"
               frmvistasprevias.Show 1
               Set reporte = Nothing
            
            
               var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
               If var_si = 6 Then
                  Set reporte = appl.OpenReport(App.Path + "\rep_kardex.rpt")
                  reporte.RecordSelectionFormula = "{VW_KARDEX.VCHA_AUD_USUARIO} = '" + vx + "' and {VW_KARDEX.VCHA_AUD_MAQUINA} = '" + vy + "' and {vw_kardex.int_kar_identificador} = " + Str(var_maximo_kardex)
                  For ntablas = 1 To reporte.Database.Tables.Count
                     reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                  Next ntablas
                  reporte.ExportOptions.FormatType = crEFTExcel80
                  reporte.ExportOptions.DestinationType = crEDTDiskFile
                  archivo = "c:\reportessid\Kardex" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                  reporte.ExportOptions.DiskFileName = archivo
                  reporte.Export False
                  Set reporte = Nothing
                  MsgBox "Se a terminado de guardar el archivo " + archivo
               End If
            
            
            Else
               MsgBox " No existen movimientos para la consulta anterior", vbOKOnly, "ATENCION"
            End If
            rs.Close
         End If

End Sub

Private Sub cmd_invertir_Click()
   Dim numero_lineas As Double
   Dim numero_seleccionado1 As Double
   Dim numero_seleccionado2 As Double
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Double
   Dim n As Double
   Dim list_item As ListItem
   Set TB_TEM_KARDEX_ARTICULOS_MARCA = New TB_TEM_KARDEX_ARTICULOS_MARCA
         n = lv_articulos.ListItems.Count
         rs.Open "update tb_tem_kardex_articulos set char_tka_marca = '1' where char_tka_marca = '*' and vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and inte_kar_identificador = " + Str(var_maximo_kardex), cnn, adOpenDynamic, adLockOptimistic
         rs.Open "update tb_tem_kardex_articulos set char_tka_marca = '*' where char_tka_marca = ' ' and vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and inte_kar_identificador = " + Str(var_maximo_kardex), cnn, adOpenDynamic, adLockOptimistic
         rs.Open "update tb_tem_kardex_articulos set char_tka_marca = ' ' where char_tka_marca = '1' and vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and inte_kar_identificador = " + Str(var_maximo_kardex), cnn, adOpenDynamic, adLockOptimistic
         For i = 1 To n
            If lv_articulos.ListItems.Item(i).SubItems(2) = "*" Then
               lv_articulos.ListItems.Item(i).SubItems(2) = " "
               lv_articulos.ListItems.Item(i).Bold = False
               lv_articulos.ListItems.Item(i).ForeColor = &H80000012
               lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
               lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
               lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
               lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
            Else
               lv_articulos.ListItems.Item(i).SubItems(2) = "*"
               lv_articulos.ListItems.Item(i).Bold = True
               lv_articulos.ListItems.Item(i).ForeColor = &H8000&
               lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
               lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
               lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
               lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
               
            End If
         Next
         lv_articulos.Refresh

End Sub

Private Sub cmd_marcar_Click()
   Dim numero_lineas As Double
   Dim numero_seleccionado1 As Double
   Dim numero_seleccionado2 As Double
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Double
   Dim n As Double
   Dim list_item As ListItem
   Set TB_TEM_KARDEX_ARTICULOS_MARCA = New TB_TEM_KARDEX_ARTICULOS_MARCA
         ok = TB_TEM_KARDEX_ARTICULOS_MARCA.Anadir(var_maximo_kardex, vx, vy, lv_articulos.selectedItem, "*")
         i = lv_articulos.selectedItem.Index
         lv_articulos.ListItems.Item(i).SubItems(2) = "*"
         lv_articulos.ListItems.Item(i).Bold = True
         lv_articulos.ListItems.Item(i).ForeColor = &H8000&
         lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
         lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         lv_articulos.Refresh

End Sub

Private Sub cmd_ninguno_Click()
   Dim numero_lineas As Double
   Dim numero_seleccionado1 As Double
   Dim numero_seleccionado2 As Double
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Double
   Dim n As Double
   Dim list_item As ListItem
   Set TB_TEM_KARDEX_ARTICULOS_MARCA = New TB_TEM_KARDEX_ARTICULOS_MARCA
         rs.Open "update tb_tem_kardex_articulos set char_tka_marca  = ' ' where vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and inte_kar_identificador = " + Str(var_maximo_kardex), cnn, adOpenDynamic, adLockOptimistic
         n = lv_articulos.ListItems.Count
         For i = 1 To n
            lv_articulos.ListItems.Item(i).SubItems(2) = " "
            lv_articulos.ListItems.Item(i).Bold = False
            lv_articulos.ListItems.Item(i).ForeColor = &H80000012
            lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = False
            lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = False
            lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
            lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         Next
         lv_articulos.Refresh

End Sub

Private Sub cmd_nuevo_Click()
   Me.lv_articulos.ListItems.Clear
   Me.txt_codigo.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub cmd_seleccion_Click()
   Dim numero_lineas As Integer
   Dim numero_seleccionado1 As Integer
   Dim numero_seleccionado2 As Integer
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Integer
   Dim n As Integer
   Dim list_item As ListItem
   Set TB_TEM_KARDEX_ARTICULOS_MARCA = New TB_TEM_KARDEX_ARTICULOS_MARCA
         primera_vez = False
         segunda_vez = False
         n = lv_articulos.ListItems.Count
         For i = 1 To n
            If lv_articulos.ListItems.Item(i).SubItems(2) = "*" And primera_vez = False Then
               numero_seleccionado1 = i
               primera_vez = True
            End If
            If lv_articulos.ListItems.Item(i).SubItems(2) = "*" And primera_vez = True Then
               numero_seleccionado2 = i
            End If
         Next
         For i = numero_seleccionado1 To numero_seleccionado2
            ok = TB_TEM_KARDEX_ARTICULOS_MARCA.Anadir(var_maximo_kardex, vx, vy, lv_articulos.ListItems.Item(i), "*")
            lv_articulos.ListItems.Item(i).SubItems(2) = "*"
            lv_articulos.ListItems.Item(i).Bold = True
            lv_articulos.ListItems.Item(i).ForeColor = &H8000&
            lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
            lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
            lv_articulos.Refresh
         Next

End Sub

Private Sub cmd_todos_Click()
   Dim numero_lineas As Double
   Dim numero_seleccionado1 As Double
   Dim numero_seleccionado2 As Double
   Dim primera_vez As Boolean
   Dim segunda_vez As Boolean
   Dim i As Double
   Dim n As Double
   Dim list_item As ListItem
   Set TB_TEM_KARDEX_ARTICULOS_MARCA = New TB_TEM_KARDEX_ARTICULOS_MARCA
         n = lv_articulos.ListItems.Count
         rs.Open "update tb_tem_kardex_articulos set char_tka_marca = '*' where vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and inte_kar_identificador = " + Str(var_maximo_kardex), cnn, adOpenDynamic, adLockOptimistic
         For i = 1 To n
            lv_articulos.ListItems.Item(i).SubItems(2) = "*"
            lv_articulos.ListItems.Item(i).Bold = True
            lv_articulos.ListItems.Item(i).ForeColor = &H8000&
            lv_articulos.ListItems.Item(i).ListSubItems(1).Bold = True
            lv_articulos.ListItems.Item(i).ListSubItems(2).Bold = True
            lv_articulos.ListItems.Item(i).ListSubItems(1).ForeColor = &H8000&
            lv_articulos.ListItems.Item(i).ListSubItems(2).ForeColor = &H8000&
         Next
         lv_articulos.Refresh

End Sub

Private Sub Form_Load()
   cnn.CommandTimeout = 600
   var_cadena_seguridad = ""
   Top = 500
   Left = 2800
   Dim cmd As New Command
   var_primera_vez = True
   vx = var_clave_usuario_global
   vy = fun_NombrePc
   rs.Open "SELECT * FROM VW_MAXIMO_KARDEX WHERE VCHA_AUD_USUARIO = '" + vx + "' AND VCHA_AUD_MAQUINA ='" + vy + "'", cnn, adOpenDynamic, adLockOptimistic
   If var_primera_vez = True Then
      var_primera_vez = False
      If Not rs.EOF Then
         var_maximo_kardex = IIf(IsNull(rs!maximo), 1, rs!maximo + 1)
      Else
         var_maximo_kardex = 1
      End If
   End If
   rs.Close
   Set cmd.ActiveConnection = cnn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "TEM_KARDEX_ARTICULOS_2"
   cmd("@MAXIMO_KARDEX") = var_maximo_kardex
   cmd("@USUARIO") = vx
   cmd("@MAQUINA") = vy
   cmd.execute
   Set cmd = Nothing
   
   'rs.Open "select * from TB_TEM_KARDEX_ARTICULOS where vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and inte_kar_identificador = " + Str(var_maximo_kardex) + " order by vcha_art_articulo_id", cnn, adOpenDynamic, adLockOptimistic
   'var_i = 0
   'While Not rs.EOF
   '    var_i = var_i + 1
   '    Set list_item = lv_articulos.ListItems.Add(, , rs(0).Value)
   '    list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
   '    list_item.SubItems(2) = " "
   '    rs.MoveNext
   'Wend
   'rs.Close
   
   rs.Open "select * from tb_almacenes where vcha_emp_empresa_id = '" + var_empresa + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
   Call RecsetToCombo(cmb_almacen.hwnd, rs, 3)
   rs.Close
   mon_mes1.Visible = False
   mon_mes2.Visible = False
   txt_inicio = Date
   txt_fin = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   cnn.CommandTimeout = 600
   rs.Open "DELETE from tb_tem_kardex where vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and INT_KAR_IDENTIFICADOR = " + Str(var_maximo_kardex), cnn, adOpenDynamic, adLockOptimistic
   rs.Open "DELETE from tb_tem_kardex_Articulos where vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and INTE_KAR_IDENTIFICADOR = " + Str(var_maximo_kardex), cnn, adOpenDynamic, adLockOptimistic
   Call activa_forma(var_activa_forma_kardex)
End Sub

Private Sub mon_mes1_DateDblClick(ByVal DateDblClicked As Date)
   txt_inicio = mon_mes1.Value
   mon_mes1.Visible = False
End Sub

Private Sub mon_mes1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mon_mes1.Visible = False
   End If
End Sub

Private Sub mon_mes2_DateDblClick(ByVal DateDblClicked As Date)
   txt_fin = mon_mes2.Value
   mon_mes2.Visible = False
End Sub

Private Sub mon_mes2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      mon_mes2.Visible = False
   End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
      Case 2
      Case 3
         Unload Me
   End Select
End Sub


Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
   If IsDate(Me.txt_inicio) Then
      mon_mes1.Value = CDate(txt_inicio)
   Else
      mon_mes1.Value = Date
   End If
   mon_mes1.Visible = True
   mon_mes1.SetFocus
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
   If IsDate(Me.txt_fin) Then
      Me.mon_mes2.Value = CDate(txt_fin)
   Else
      Me.mon_mes2.Value = Date
   End If
   mon_mes2.Visible = True
   mon_mes2.SetFocus
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Dim list_item As ListItem
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      rs.Open "select * from tb_articulos where vcha_art_articulo_id = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.EOF Then
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            txt_codigo = rsaux!vcha_Art_articulo_id
         Else
            txt_codigo = ""
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rsaux.Close
      End If
      rs.Close
      If Len(Trim(txt_codigo)) > 0 Then
         Set TB_TEM_KARDEX_ARTICULOS_MARCA = New TB_TEM_KARDEX_ARTICULOS_MARCA
         rs.Open "select * from tb_Articulos where vcha_art_articulo_id ='" + txt_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            rsaux.Open "select * from TB_TEM_KARDEX_ARTICULOS where inte_kar_identificador = " + CStr(var_maximo_kardex) + " and vcha_aud_usuario = '" + vx + "' and vcha_aud_maquina = '" + vy + "' and vcha_Art_articulo_id = '" + txt_codigo + "' and char_tka_marca = '*'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               txt_descripcion = rs!vcha_art_nombre_español
               valor = txt_codigo
               Set itmfound = lv_articulos.findItem(valor, lvwText, , lvwPartial)
               itmfound.EnsureVisible
               itmfound.Selected = True
               ok = TB_TEM_KARDEX_ARTICULOS_MARCA.Anadir(var_maximo_kardex, vx, vy, lv_articulos.selectedItem, "*")
               lv_articulos.Refresh
            Else
               txt_descripcion = rs!vcha_art_nombre_español
               valor = txt_codigo
               Set list_item = lv_articulos.ListItems.Add(, , txt_codigo)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_art_nombre_español), "", rs!vcha_art_nombre_español)
               ok = TB_TEM_KARDEX_ARTICULOS_MARCA.Anadir(var_maximo_kardex, vx, vy, txt_codigo, "*")
            End If
         Else
            rsaux.Close
            txt_descripcion = ""
            MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
      Me.txt_codigo = ""
      Me.txt_descripcion = ""
   End If
End Sub
