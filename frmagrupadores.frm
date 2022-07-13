VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmagrupadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agrupadores"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmagrupadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.CommandButton cmd_clonar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmagrupadores.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Clonar Agrupador Alt + L"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmagrupadores.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmagrupadores.frx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmagrupadores.frx":0BD0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmagrupadores.frx":0CA2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      Picture         =   "frmagrupadores.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmagrupadores.frx":0EA6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2190
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Agrupadores "
      Height          =   2895
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_texto 
         Height          =   1485
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   5370
      End
      Begin VB.TextBox txt_pais 
         Height          =   315
         Left            =   4455
         MaxLength       =   3
         TabIndex        =   12
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox txt_fraccion 
         Height          =   315
         Left            =   2895
         MaxLength       =   50
         TabIndex        =   11
         Top             =   930
         Width           =   900
      End
      Begin VB.TextBox txt_agrupadores 
         Height          =   315
         Index           =   2
         Left            =   765
         MaxLength       =   1
         TabIndex        =   10
         Top             =   915
         Width           =   405
      End
      Begin VB.TextBox txt_agrupadores 
         Height          =   315
         Index           =   1
         Left            =   765
         MaxLength       =   50
         TabIndex        =   9
         Top             =   585
         Width           =   4815
      End
      Begin VB.TextBox txt_agrupadores 
         Height          =   315
         Index           =   0
         Left            =   765
         MaxLength       =   50
         TabIndex        =   8
         Top             =   255
         Width           =   900
      End
      Begin MSComctlLib.Toolbar tool_grupos 
         Height          =   330
         Left            =   5250
         TabIndex        =   14
         Top             =   945
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Detalle de agrupador"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Left            =   4050
         TabIndex        =   27
         Top             =   975
         Width           =   345
      End
      Begin VB.Label Label2 
         Caption         =   "Fracción arancelaria:"
         Height          =   240
         Left            =   1395
         TabIndex        =   26
         Top             =   975
         Width           =   1560
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   20
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   615
         Width           =   600
      End
      Begin VB.Label lab_estados 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   135
      TabIndex        =   21
      Top             =   3345
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   1965
         TabIndex        =   15
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   3915
         TabIndex        =   16
         Top             =   165
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al primero"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un Registro Atras"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Un registro adelante"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ir al ultimo"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda de agrupador:"
         Height          =   195
         Left            =   195
         TabIndex        =   22
         Top             =   195
         Width           =   1755
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3420
      Left            =   150
      TabIndex        =   23
      Top             =   3855
      Width           =   5655
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   60
         Top             =   -315
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   12
         ImageHeight     =   12
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupadores.frx":14E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmagrupadores.frx":1DBA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_agrupadores 
         Height          =   3210
         Left            =   45
         TabIndex        =   17
         Top             =   165
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   5662
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ancho"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "fraccion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "pais"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "texto"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   24
      Top             =   285
      Width           =   5655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -75
      Top             =   2205
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":2694
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":2F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":3848
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":3DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":4F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":5874
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":5986
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":5A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":5BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmagrupadores.frx":5CBC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmagrupadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_agrupadores As Integer
Dim bitacora As Boolean
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report





Private Sub cmd_clonar_Click()
   frmclonacionagrupadores.Show
   lv_agrupadores.ListItems.Clear
   Call pro_llena_listview1
End Sub

Private Sub cmd_deshacer_Click()
      Call pro_textos

End Sub

Private Sub cmd_eliminar_Click()
   var_opcion_seguridad = 2
   var_acepta_seguridad = 1
   If var_global_permiso3 = 1 Then
      var_acepta_seguridad = 2
      If var_global_permiso4 = 1 Then
         frmpasswords2.Show 1
      Else
         frmpasswords.Show 1
      End If
   End If
   
   If var_acepta_seguridad = 1 Then
      Call pro_elimina_agrupadores
      rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + varfamiliaagrupadores + "'", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
      End If
      rs.Close
   End If
End Sub

Private Sub cmd_guardar_Click()
Dim var_posible As Boolean
   var_posible = True
   If var_modifica_registro_agrupadores = False Then
      rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + Me.txt_agrupadores(0) + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_posible = False
      End If
      rs.Close
   End If
   If var_posible = True Then
      var_opcion_seguridad = 2
      var_acepta_seguridad = 1
      If var_global_permiso3 = 1 Then
         var_acepta_seguridad = 2
         If var_global_permiso4 = 1 Then
            frmpasswords2.Show 1
         Else
            frmpasswords.Show 1
         End If
      End If
      If var_acepta_seguridad = 1 Then
         Call pro_guardar_agrupadores
         rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + varfamiliaagrupadores + "'", cnn, adOpenDynamic, adLockOptimistic
         If rs.BOF Then
            cmd_guardar.Enabled = False
            cmd_deshacer.Enabled = False
            cmd_eliminar.Enabled = False
         Else
            cmd_guardar.Enabled = True
            cmd_deshacer.Enabled = True
            cmd_eliminar.Enabled = True
         End If
         rs.Close
      End If
   Else
      MsgBox "Clave de familia de agrupador ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Set reporte = appl.OpenReport(App.Path + "\rep_agrupadores.rpt")
   reporte.RecordSelectionFormula = "{VW_AGRUPADOR_ARTICULOS.VCHA_AGR_AGRUPADOR_ID} = '" + Me.txt_agrupadores(0).Text + "'"
   frmvistasprevias.cr.ReportSource = reporte
   For ntablas = 1 To reporte.Database.Tables.Count
       reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
   Next ntablas
   frmvistasprevias.cr.ViewReport
   frmvistasprevias.Caption = "Reporte de Entradas concentrado"
   frmvistasprevias.Show 1
   Set reporte = Nothing
   var_si = MsgBox("¿Desea importar el reporte?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      Set reporte = appl.OpenReport(App.Path + "\rep_agrupadores.rpt")
      reporte.RecordSelectionFormula = "{VW_AGRUPADOR_ARTICULOS.VCHA_AGR_AGRUPADOR_ID} = '" + Me.txt_agrupadores(0).Text + "'"
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      reporte.ExportOptions.FormatType = crEFTExcel80
      reporte.ExportOptions.DestinationType = crEDTDiskFile
      archivo = "c:\reportessid\REPORTE_AGRUPADORES_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
      reporte.ExportOptions.DiskFileName = archivo
      reporte.Export False
      Set reporte = Nothing
      MsgBox "Se a terminado de guardar el archivo " + archivo
   End If
End Sub

Private Sub cmd_nuevo_Click()

      Call pro_limpiatextos(Me)
      rsaux9.Open "select max(cast(vcha_agr_agrupador_id as integer)) from tb_agrupadores where cast(vcha_agr_agrupador_id as integer) < 50000 and vcha_agr_agrupador_id <> 'faltan'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux9.EOF Then
         Me.txt_agrupadores(0).Text = rsaux9(0).Value + 1
      End If
      rsaux9.Close
      txt_agrupadores(0).Enabled = True
      
      txt_agrupadores(0).SetFocus: var_modifica_registro_agrupadores = False
      Me.txt_agrupadores(1).SetFocus
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True

End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_agrupadores = False Then
      var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
      If var_si <> 6 Then
         GoTo salir:
      End If
   Else
      If var_hubo_cambios = True Then
         var_si = MsgBox("No se han guardado los cambios, ¿Desea salir?", vbYesNo, "ATENCION")
         If var_si <> 6 Then
            GoTo salir:
         End If
      End If
   End If
   var_swpassword = False
   var_modifica_registro_agrupadores = False
   Unload Me
   Exit Sub
salir:
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
      cmd_deshacer_Click
   End If
   If Shift = 4 And KeyCode = 69 Then
      cmd_eliminar_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
   If Shift = 4 And KeyCode = 76 Then
      cmd_clonar_Click
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   varfamiliaagrupadores = frmfamilia_agrupadores.txt_familia_agrupadores(0)
   var_modifica_registro_agrupadores = True
   lv_agrupadores.SmallIcons = ImageList1
   Call pro_encabezadosView(Me, lv_agrupadores, False)
   Call pro_llena_listview1
   pro_textos
   rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + varfamiliaagrupadores + "'", cnn, adOpenDynamic, adLockOptimistic
   If rs.BOF Then
      cmd_guardar.Enabled = False
      cmd_deshacer.Enabled = False
      cmd_eliminar.Enabled = False
   Else
      cmd_guardar.Enabled = True
      cmd_deshacer.Enabled = True
      cmd_eliminar.Enabled = True
   End If
   rs.Close
   Me.txt_agrupadores(0).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_agrupadores)
End Sub

Private Sub lv_agrupadores_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_agrupadores, ColumnHeader)
End Sub

Private Sub lv_agrupadores_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_agrupadores.selectedItem = Item
        pro_textos
        var_modifica_registro_agrupadores = True
        txt_agrupadores(0).Enabled = False

End Sub



Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_agrupadores.SetFocus
      Call pro_avanzar(Me, lv_agrupadores, Button)
      lv_agrupadores.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_agrupadores.ListItems(1).Selected = True
      lv_agrupadores.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_agrupadores = lv_agrupadores.ListItems.Count
      lv_agrupadores.ListItems(numero_items_agrupadores).Selected = True
      lv_agrupadores.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub

Private Sub tool_grupos_ButtonClick(ByVal Button As MSComctlLib.Button)
   If lv_agrupadores.ListItems.Count > 0 Then
      frmdetalleagrupadores.Caption = "Agrupador:  " + lv_agrupadores.selectedItem.SubItems(1)
      frmdetalleagrupadores.Show
   Else
      MsgBox "No existe un agrupador", vbOKOnly, "ATENCION"
   End If
End Sub


Sub pro_guardar_agrupadores()

Dim ok As Boolean

Set TB_AGRUPADORES = New TB_AGRUPADORES
Set TB_BITACORA_AGRUPADORES = New TB_BITACORA_AGRUPADORES
    
    ok = True
    If txt_agrupadores(0) <> "" And txt_agrupadores(1) <> "" Then
        If var_hubo_cambios Then
           rs.Open "select * from tb_agrupadores where VCHA_FAG_FAMILIA_AGRUPADOR_ID = '" + varfamiliaagrupadores + "' and VCHA_AGR_AGRUPADOR_ID = '" + txt_agrupadores(0) + "'", cnn, adOpenDynamic, adLockOptimistic
           If Trim(Me.txt_fraccion) = "" Then
              Me.txt_fraccion = "0"
           End If
           If Me.txt_agrupadores(2) = "" Then
              Me.txt_agrupadores(2) = 1
           End If
           ok = TB_AGRUPADORES.Anadir(varfamiliaagrupadores, txt_agrupadores(0), txt_agrupadores(1), txt_agrupadores(2), CStr(Me.txt_fraccion), Me.txt_pais, Me.txt_texto)
           If ok Then
              bitacora = True
              If var_modifica_registro_agrupadores = False Then
                 var_operacion_bitacora = "I"
                 bitacora = TB_BITACORA_AGRUPADORES.Anadir(varfamiliaagrupadores, txt_agrupadores(0), "VCHA_AGR_NOMBRE", txt_agrupadores(1), "", var_clave_usuario_global, fun_NombrePc, Date)
              Else
                 var_operacion_bitacora = "M"
                 If rs(1) <> txt_agrupadores(0) Then
                    bitacora = TB_BITACORA_AGRUPADORES.Anadir(varfamiliaagrupadores, txt_agrupadores(0), "VCHA_AGR_AGRUPADOR_ID", rs(0), txt_agrupadores(0), var_clave_usuario_global, fun_NombrePc, Date)
                 End If
                 If rs(2) <> txt_agrupadores(1) Then
                    bitacora = TB_BITACORA_AGRUPADORES.Anadir(varfamiliaagrupadores, txt_agrupadores(0), "VCHA_AGR_NOMBRE", rs(2), txt_agrupadores(1), var_clave_usuario_global, fun_NombrePc, Date)
                 End If
                 If rs(3) <> txt_agrupadores(2) Then
                    bitacora = TB_BITACORA_AGRUPADORES.Anadir(varfamiliaagrupadores, txt_agrupadores(0), "VCHA_AGR_TIPO", rs(3), txt_agrupadores(2), var_clave_usuario_global, fun_NombrePc, Date)
                 End If
              End If
              rs.Close
              pro_actualiza_ListView
              txt_agrupadores(0).Enabled = False
              MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
              txt_registros = lv_agrupadores.ListItems.Count
              var_modifica_registro_agrupadores = True
          Else
             rs.Close
             MsgBox "No se puede grabar registro: " + TB_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
          End If
      End If
  End If
    
Set TB_AGRUPADORES = Nothing: var_hubo_cambios = False

End Sub

Sub pro_elimina_agrupadores()
   Dim var_llave_usuarios As String
   Set TB_AGRUPADORES = New TB_AGRUPADORES
   Set TB_BITACORA_AGRUPADORES = New TB_BITACORA_AGRUPADORES
   
   ok = True
   On Error GoTo salir:
   If txt_agrupadores(0) <> "" And txt_agrupadores(1) <> "" And var_modifica_registro_agrupadores = True Then
      If MsgBox("Desea Eliminar este Registro", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_AGRUPADORES.Eliminar(txt_agrupadores(0))
      Else
         GoTo salir:
      End If
      If ok Then
         bitacora = True
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_AGRUPADORES.Anadir(varfamiliaagrupadores, txt_agrupadores(0), "", txt_agrupadores(1), "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_agrupadores = numero_items_agrupadores - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_agrupadores.ListItems.Remove (lv_agrupadores.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_agrupadores.ListItems.Count
         lv_agrupadores.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede eliminar registro: " + TB_AGRUPADORES.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_AGRUPADORES = Nothing
End Sub


Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from tb_agrupadores where vcha_fag_familia_agrupador_id = '" + varfamiliaagrupadores + "'", cnn, adOpenDynamic, adLockOptimistic
   numero_items_agrupadores = 0
   While Not rs.EOF
      Set list_item = lv_agrupadores.ListItems.Add(, , rs(1).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(2).Value), "", rs(2).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(3).Value), "", rs(3).Value)
      list_item.SubItems(3) = IIf(IsNull(rs!floa_agr_fraccion_arancelaria), "", rs!floa_agr_fraccion_arancelaria)
      list_item.SubItems(4) = IIf(IsNull(rs!vcha_agr_pais), "", rs!vcha_agr_pais)
      list_item.SubItems(5) = IIf(IsNull(rs!vcha_agr_texto), "", rs!vcha_agr_texto)
      rs.MoveNext:
      numero_items_agrupadores = numero_items_agrupadores + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Double
var_n = lv_agrupadores.ListItems.Count
   If var_n > 0 Then
      txt_agrupadores(0) = lv_agrupadores.selectedItem
      txt_agrupadores(1) = lv_agrupadores.selectedItem.SubItems(1)
      txt_agrupadores(2) = lv_agrupadores.selectedItem.SubItems(2)
      Me.txt_fraccion = lv_agrupadores.selectedItem.SubItems(3)
      Me.txt_pais = lv_agrupadores.selectedItem.SubItems(4)
      Me.txt_texto = lv_agrupadores.selectedItem.SubItems(5)
   End If
   var_numero_renglones = lv_agrupadores.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_agrupadores.ColumnHeaders(2).Width = 3850
   Else
      lv_agrupadores.ColumnHeaders(2).Width = 4099.9
   End If
   var_hubo_cambios = False
   var_modifica_registro_agrupadores = True
   Me.txt_agrupadores(0).Enabled = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem

    If var_modifica_registro_agrupadores = False Then
        Set list_item = lv_agrupadores.ListItems.Add(, , txt_agrupadores(0))
        list_item.SubItems(1) = txt_agrupadores(1)
        list_item.SubItems(2) = txt_agrupadores(2)
        list_item.SubItems(3) = Me.txt_fraccion
        list_item.SubItems(4) = Me.txt_pais
        list_item.SubItems(5) = Me.txt_texto
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_agrupadores = numero_items_agrupadores + 1
    Else
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).Checked = False
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index) = txt_agrupadores(0)
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).ListSubItems(1) = txt_agrupadores(1)
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).ListSubItems(2) = txt_agrupadores(2)
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).ListSubItems(3) = Me.txt_fraccion
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).ListSubItems(4) = Me.txt_pais
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).ListSubItems(5) = Me.txt_texto
        lv_agrupadores.ListItems.Item(lv_agrupadores.selectedItem.Index).Selected = True
    End If
'    lv_agrupadores.SetFocus
End Sub

Private Sub txt_agrupadores_Change(Index As Integer)
   var_hubo_cambios = True
End Sub

Private Sub txt_agrupadores_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_agrupadores, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_fraccion_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fraccion_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_pais_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

