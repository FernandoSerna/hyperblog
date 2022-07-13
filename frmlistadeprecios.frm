VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmlistadeprecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de precios"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmlistadeprecios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5955
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   135
      TabIndex        =   31
      Top             =   1335
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   120
         Width           =   5610
      End
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Index           =   1
      Left            =   2190
      TabIndex        =   28
      Top             =   1140
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   69337089
      CurrentDate     =   37581
   End
   Begin MSComCtl2.MonthView mes 
      Height          =   2370
      Index           =   0
      Left            =   2205
      TabIndex        =   29
      Top             =   1110
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   69337089
      CurrentDate     =   37581
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Picture         =   "frmlistadeprecios.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Detalle de la Lista de precios Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5460
      Picture         =   "frmlistadeprecios.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      Picture         =   "frmlistadeprecios.frx":1006
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cargar precios"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Picture         =   "frmlistadeprecios.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   810
      Picture         =   "frmlistadeprecios.frx":120A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmlistadeprecios.frx":12DC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar Alt + G"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmlistadeprecios.frx":13DE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   30
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Lista de precios "
      Height          =   1995
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   5655
      Begin VB.TextBox txt_nombre_moneda 
         Height          =   315
         Left            =   2175
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1575
         Width           =   3390
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   1
         Left            =   2475
         Picture         =   "frmlistadeprecios.frx":14E0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Seleccione la fecha"
         Top             =   1245
         Width           =   315
      End
      Begin VB.CommandButton cmdfecha 
         Height          =   285
         Index           =   0
         Left            =   2475
         Picture         =   "frmlistadeprecios.frx":15E2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Seleccione la fecha"
         Top             =   945
         Width           =   315
      End
      Begin VB.TextBox txt_moneda 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1575
         Width           =   900
      End
      Begin VB.TextBox txt_fecha_fin 
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1245
         Width           =   1185
      End
      Begin VB.TextBox txt_fecha_inicio 
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   915
         Width           =   1185
      End
      Begin VB.TextBox txt_nombre_lista_precios 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   9
         Top             =   585
         Width           =   4290
      End
      Begin VB.TextBox txt_lista_precios 
         Height          =   315
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   8
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   1620
         Width           =   630
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha final:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   1305
         Width           =   825
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de inicio:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   150
      TabIndex        =   19
      Top             =   2460
      Width           =   5655
      Begin VB.TextBox txt_buscar 
         Height          =   315
         Left            =   2430
         TabIndex        =   30
         Top             =   150
         Width           =   1350
      End
      Begin MSComctlLib.Toolbar tool_atras_siguiente 
         Height          =   330
         Left            =   4035
         TabIndex        =   27
         Top             =   150
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList"
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
         Caption         =   "Busqueda de lista de precios:"
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   195
         Width           =   2085
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4230
      Left            =   150
      TabIndex        =   21
      Top             =   2955
      Width           =   5655
      Begin MSComctlLib.ListView lv_listadeprecios 
         Height          =   4035
         Left            =   45
         TabIndex        =   26
         Top             =   135
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   7117
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6879
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "fecha_inicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "fecha_fin"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "moneda"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList icono_encabezado 
      Left            =   2385
      Top             =   -45
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
            Picture         =   "frmlistadeprecios.frx":16E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":1FBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   150
      TabIndex        =   22
      Top             =   285
      Width           =   5655
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmlistadeprecios.frx":2898
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":3172
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":3A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":3FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":48C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":519E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":5A78
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":5B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":5C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":5DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlistadeprecios.frx":5EC0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmlistadeprecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_hubo_cambios As Boolean
Dim numero_items_listadeprecios As Integer
Dim bitacora As Boolean








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
      Call pro_elimina_listadeprecios
      rs.Open "select * from tb_listadeprecios", cnn, adOpenDynamic, adLockOptimistic
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
   If var_modifica_registro_lista_precios = False Then
      rs.Open "SELECT * FROM TB_LISTADEPRECIOS WHERE vcha_lis_lista_id = '" + Me.txt_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
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
         Call pro_guardar_listadeprecios
         rs.Open "update tb_listadeprecios set vcha_Emp_empresa_id = '" + var_empresa + "'  WHERE vcha_lis_lista_id = '" + Me.txt_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
         rs.Open "select * from tb_listadeprecios ", cnn, adOpenDynamic, adLockOptimistic
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
      MsgBox "Clave de lista de precios ya existe", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim strConnectionString As String
   On Error GoTo salir:
   var_cadena = App.Path + "\lista_precios.xls"
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & App.Path + "\lista_precios.xls"
   rsaux2.Open "SELECT LISTA, CODIGO, PRECIO, CATALOGO FROM [lista$] WHERE LISTA IS NOT NULL AND CODIGO IS NOT NULL", strConnectionString, adOpenDynamic, adLockOptimistic
   If Not rsaux2.EOF Then
      var_si = MsgBox("¿Desea actualizar la lista de precios?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la actualizacion de la lista de precios", vbYesNo, "ATENCION")
         If var_si = 6 Then
            var_lista_precios = IIf(IsNull(rsaux2!LISTA), "", rsaux2!LISTA)
            rsaux4.Open "SELECT * FROM TB_LISTADEPRECIOS WHERE VCHA_LIS_LISTA_ID = '" + var_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux4.EOF Then
               cnn.BeginTrans
               rsaux5.Open "SELECT MAX(INTE_BIT_NUMERO_RESPALDO) FROM TB_BITACORA_LISTA_PRECIOS", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux5.EOF Then
                  var_consecutivo = IIf(IsNull(rsaux5(0).Value), 0, rsaux5(0).Value) + 1
               Else
                  var_consecutivo = 1
               End If
               rsaux5.Close
               var_cadena = " INSERT INTO TB_BITACORA_LISTA_PRECIOS (VCHA_BIT_USUARIO, VCHA_BIT_MAQUINA,VCHA_LIS_LISTA_PRECIOS_ID, VCHA_ART_ARTICULO_ID, FLOA_BIT_PRECIO_ANTERIOR, FLOA_BIT_PRECIO_ACTUAL,VCHA_BIT_CATALOGO_ANTERIOR, VCHA_BIT_CATALOGO_ACTUAL, VCHA_BIT_OPERACION, INTE_BIT_NUMERO_RESPALDO)"
               var_cadena = var_cadena + " select '" + var_clave_usuario_global + "', '" + fun_NombrePc + "', vcha_lis_lista_precios_id, VCHA_ART_ARTICULO_ID, floa_dli_precio,0, vcha_Cat_catalogo_id, '',''," + CStr(var_consecutivo) + " FROM TB_DETALLE_LISTA_PRECIOS WHERE VCHA_LIS_LISTA_PRECIOS_ID = '" + rsaux2!LISTA + "'"
               rsaux3.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
               cnn.CommitTrans
               While Not rsaux2.EOF
                     rsaux5.Open "UPDATE TB_DETALLE_LISTA_PRECIOS SET VCHA_CAT_CATALOGO_ID = '" + IIf(IsNull(rsaux2!CATALOGO), "", rsaux2!CATALOGO) + "' WHERE VCHA_LIS_LISTA_PRECIOS_ID = '" + IIf(IsNull(rsaux2!LISTA), "", rsaux2!LISTA) + "' AND VCHA_ART_ARTICULO_ID = '" + IIf(IsNull(rsaux2!codigo), "", rsaux2!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                     rsaux5.Open "UPDATE TB_BITACORA_LISTA_PRECIOS SET VCHA_BIT_CATALOGO_ACTUAL = '" + IIf(IsNull(rsaux2!CATALOGO), "", rsaux2!CATALOGO) + "'  WHERE VCHA_LIS_LISTA_PRECIOS_ID = '" + IIf(IsNull(rsaux2!LISTA), "", rsaux2!LISTA) + "' AND VCHA_ART_ARTICULO_ID = '" + IIf(IsNull(rsaux2!codigo), "", rsaux2!codigo) + "' AND INTE_BIT_NUMERO_RESPALDO = " + CStr(var_consecutivo)
                     rsaux2.MoveNext
               Wend
               MsgBox "Se a terminado de actualizar la lista de precios", vbOKOnly, "ATENCION"
            Else
               MsgBox "La lista de precios no existe", vbOKOnly, "ATENCION"
            End If
            rsaux4.Close
         End If
      End If
   Else
      MsgBox "El archivo se encuentra vacio", vbOKOnly, "ATENCION"
   End If
   rsaux2.Close
   Exit Sub
salir:
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
   MsgBox "Error al cargar el archivo, puede que este no exista o no tenga el formato adecuado que debe de ser LISTA, CODIGO, PRECIO, CATALOGO", vbOKOnly, "ATENCION"
End Sub

Private Sub cmd_nuevo_Click()
        Call pro_limpiatextos(Me)
        txt_lista_precios.Enabled = True
        txt_lista_precios.SetFocus: var_modifica_registro_lista_precios = False
        cmd_guardar.Enabled = True
        cmd_deshacer.Enabled = True
End Sub

Private Sub cmd_salir_Click()
   Dim var_si As Integer
   If var_modifica_registro_lista_precios = False Then
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
   Unload Me
   Exit Sub
salir:
End Sub

Private Sub cmdfecha_Click(Index As Integer)
   If Index = 0 Then
      If IsDate(Me.txt_fecha_inicio) Then
         mes(0) = CDate(Me.txt_fecha_inicio)
      Else
         mes(0) = Date
      End If
      mes(0).Visible = True
      mes(0).SetFocus
   End If
   If Index = 1 Then
      If IsDate(Me.txt_fecha_fin) Then
         mes(1) = CDate(Me.txt_fecha_fin)
      Else
         mes(1) = Date
      End If
      mes(1).Visible = True
      mes(1).SetFocus
   End If
End Sub


Private Sub Command1_Click()
   var_clave_lista = txt_lista_precios
   frmdetalle_lista_precios.txt_lista = txt_lista_precios
   frmdetalle_lista_precios.Caption = "DETALLE DE LA LISTA: " + Me.txt_nombre_lista_precios
   frmdetalle_lista_precios.Show
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
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 2900
   frm_lista.Visible = False
   var_modifica_registro_lista_precios = True
   lv_listadeprecios.SmallIcons = ImageList
   Call pro_encabezadosView(Me, lv_listadeprecios, False)
   Call pro_llena_listview1
   pro_textos
   If VAR_EMPRESA_ID = "16" Then
      rs.Open "select * from tb_listadeprecios WHERE VCHA_EMP_EMPRESA_ID = '16'", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
         txt_lista_precios.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
         txt_lista_precios.Enabled = True
      End If
      rs.Close
   Else
      rs.Open "select * from tb_listadeprecios WHERE VCHA_EMP_EMPRESA_ID <> '16' OR VCHA_EMP_EMPRESA_ID IS NULL", cnn, adOpenDynamic, adLockOptimistic
      If rs.BOF Then
         cmd_guardar.Enabled = False
         cmd_deshacer.Enabled = False
         cmd_eliminar.Enabled = False
         txt_lista_precios.Enabled = False
      Else
         cmd_guardar.Enabled = True
         cmd_deshacer.Enabled = True
         cmd_eliminar.Enabled = True
         txt_lista_precios.Enabled = True
      End If
      rs.Close
   End If
   mes(1).Visible = False
   mes(0).Visible = False
   Me.txt_lista_precios.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro_lista_precios = False
   Call activa_forma(var_activa_forma_listadeprecios)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_moneda = lv_lista.selectedItem
         txt_nombre_moneda = lv_lista.selectedItem.SubItems(1)
      Else
         txt_moneda = ""
         txt_nombre_moneda = ""
      End If
      txt_moneda.SetFocus
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub lv_listadeprecios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_listadeprecios, ColumnHeader)
End Sub

Private Sub lv_listadeprecios_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lv_listadeprecios.selectedItem = Item
        pro_textos
        var_modifica_registro_lista_precios = True
        txt_lista_precios.Enabled = False
End Sub

Private Sub mes_DateDblClick(Index As Integer, ByVal DateDblClicked As Date)
   If Index = 0 Then
      txt_fecha_inicio = mes(0).Value
      mes(0).Visible = False
      Me.txt_fecha_fin.SetFocus
   End If
   If Index = 1 Then
      txt_fecha_fin = mes(1).Value
      mes(1).Visible = False
      Me.txt_moneda.SetFocus
   End If
End Sub

Private Sub mes_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      mes(0).Visible = False
      mes(1).Visible = False
   End If
End Sub

Private Sub mes_LostFocus(Index As Integer)
   mes(0).Visible = False
   mes(1).Visible = False
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_listadeprecios.SetFocus
      Call pro_avanzar(Me, lv_listadeprecios, Button)
      Me.lv_listadeprecios.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_listadeprecios.ListItems(1).Selected = True
      Me.lv_listadeprecios.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_listadeprecios = Me.lv_listadeprecios.ListItems.Count
      lv_listadeprecios.ListItems(numero_items_listadeprecios).Selected = True
      Me.lv_listadeprecios.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub


Sub pro_guardar_listadeprecios()
Dim ok As Boolean
   Dim var_clave As String
   Dim var_nombre As String
   Dim var_fecha_inicio As Date
   Dim var_fecha_fin As Date
   Dim var_moneda As String
   Set TB_LISTADEPRECIOS = New TB_LISTADEPRECIOS
   Set TB_BITACORA_LISTA_PECIOS = New TB_BITACORA_LISTA_PECIOS
   If txt_lista_precios <> "" And txt_nombre_lista_precios <> "" Then
      If var_hubo_cambios Then
         If Not IsDate(Me.txt_fecha_fin) Then
            Me.txt_fecha_fin = Date + 10000
         End If
         If Not IsDate(Me.txt_fecha_inicio) Then
            Me.txt_fecha_inicio = Date
         End If
         rs.Open "select * from tb_listadeprecios where vcha_lis_lista_id = '" + txt_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
         ok = TB_LISTADEPRECIOS.Anadir(txt_lista_precios, txt_nombre_lista_precios, txt_fecha_inicio, txt_fecha_fin, txt_moneda)
         If ok Then
            bitacora = True
            If var_modifica_registro_lista_precios = False Then
               var_operacion_bitacora = "I"
               bitacora = TB_BITACORA_LISTA_PECIOS.Anadir(txt_lista_precios, "VCHA_LIS_NOMBRE", var_operacion_bitacora, "", txt_nombre_lista_precios, var_clave_usuario_global, fun_NombrePc, Date)
            Else
               var_operacion_bitacora = "M"
               If rs(0) <> txt_lista_precios Then
                  bitacora = TB_BITACORA_LISTA_PECIOS.Anadir(txt_lista_precios, "VCHA_LIS_LISTA_ID", var_operacion_bitacora, rs(0), txt_lista_precios, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(1) <> txt_nombre_lista_precios Then
                  bitacora = TB_BITACORA_LISTA_PECIOS.Anadir(txt_lista_precios, "VCHA_LIS_NOMBRE", var_operacion_bitacora, rs(1), txt_nombre_lista_precios, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(2) <> txt_fecha_inicio Then
                  bitacora = TB_BITACORA_LISTA_PECIOS.Anadir(txt_lista_precios, "VCHA_LIS_FECHA_INICIO_ID", var_operacion_bitacora, rs(2), txt_fecha_inicio, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(3) <> txt_fecha_fin Then
                  bitacora = TB_BITACORA_LISTA_PECIOS.Anadir(txt_lista_precios, "VCHA_LIS_FECHA_FIN_NOMBRE", var_operacion_bitacora, rs(3), txt_fecha_fin, var_clave_usuario_global, fun_NombrePc, Date)
               End If
               If rs(4) <> txt_moneda Then
                  bitacora = TB_BITACORA_LISTA_PECIOS.Anadir(txt_lista_precios, "VCHA__MON_MONEDA_ID", var_operacion_bitacora, rs(4), txt_moneda, var_clave_usuario_global, fun_NombrePc, Date)
               End If
             End If
            rs.Close
            pro_actualiza_ListView
            txt_lista_precios.Enabled = False
            MsgBox "Informacion Guardada Correctamente ! ", vbOKOnly + vbInformation, "Aviso"
            txt_registros = lv_listadeprecios.ListItems.Count
            var_modifica_registro_lista_precios = True
         Else
            MsgBox "No se puede grabar registro: " + TB_LISTADEPRECIOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
         End If
      End If
   End If
   Set TB_LISTADEPRECIOS = Nothing: var_hubo_cambios = False
End Sub

Sub pro_elimina_listadeprecios()
   Dim var_llave_usuarios As String
   Set TB_LISTADEPRECIOS = New TB_LISTADEPRECIOS
   Set TB_BITACORA_LISTA_PECIOS = New TB_BITACORA_LISTA_PECIOS
   On Error GoTo salir
   ok = True
   If txt_lista_precios <> "" And txt_nombre_lista_precios <> "" And var_modifica_registro_lista_precios = True Then
      If MsgBox("Desea Eliminar este Registro, con esto también eliminara el detalle correspondiente", vbInformation + vbYesNo, "Aviso") = vbYes Then
         ok = TB_LISTADEPRECIOS.Eliminar(txt_lista_precios)
      Else
         GoTo salir:
      End If
      If ok Then
         rs.Open "delete from TB_DETALLE_LISTA_PRECIOS where VCHA_LIS_LISTA_PRECIOS_ID = '" + txt_lista_precios + "'", cnn, adOpenDynamic, adLockOptimistic
         var_operacion_bitacora = "E"
         bitacora = TB_BITACORA_LISTA_PECIOS.Anadir(txt_lista_precios, "VCHA_LIS_NOMBRE", var_operacion_bitacora, txt_nombre_lista_precios, "", var_clave_usuario_global, fun_NombrePc, Date)
         numero_items_listadeprecios = numero_items_listadeprecios - 1
         MsgBox "Se Elimino Correctamente el Registro", vbInformation
         lv_listadeprecios.ListItems.Remove (lv_listadeprecios.selectedItem.Index)
         Call pro_limpiatextos(Me)
         txt_registros = lv_listadeprecios.ListItems.Count
         lv_listadeprecios.selectedItem.Selected = True
         pro_textos
      Else
         MsgBox "No se puede grabar registro: " + TB_LISTADEPRECIOS.MensajeError, vbOKOnly + vbCritical, "ATENCION"
      End If
   End If
salir:
   Set TB_LISTADEPRECIOS = Nothing
End Sub


Sub pro_llena_listview1()

Dim list_item As ListItem
    If var_empresa = "16" Then
       rs.Open "select * from TB_listadeprecios WHERE VCHA_EMP_EMPRESA_ID = '16'", cnn, adOpenDynamic, adLockOptimistic
       numero_items_listadeprecios = 0
       While Not rs.EOF
           Set list_item = lv_listadeprecios.ListItems.Add(, , rs!vcha_lis_lista_id)
           list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lis_NOMBRE), "", rs!VCHA_lis_NOMBRE)
           list_item.SubItems(2) = IIf(IsNull(rs!DTIM_LIS_FECHA_INICIO), "", rs!DTIM_LIS_FECHA_INICIO)
           list_item.SubItems(3) = IIf(IsNull(rs!DTIM_LIS_FECHA_FIN), "", rs!DTIM_LIS_FECHA_FIN)
           list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_moneda), "", rs!vcha_mon_moneda)
           rs.MoveNext:
           numero_items_listadeprecios = numero_items_listadeprecios + 1
       Wend
       rs.Close
    Else
       rs.Open "select * from TB_listadeprecios WHERE VCHA_EMP_EMPRESA_ID <>'16' OR VCHA_EMP_EMPRESA_ID IS NULL", cnn, adOpenDynamic, adLockOptimistic
       numero_items_listadeprecios = 0
       While Not rs.EOF
           Set list_item = lv_listadeprecios.ListItems.Add(, , rs!vcha_lis_lista_id)
           list_item.SubItems(1) = IIf(IsNull(rs!VCHA_lis_NOMBRE), "", rs!VCHA_lis_NOMBRE)
           list_item.SubItems(2) = IIf(IsNull(rs!DTIM_LIS_FECHA_INICIO), "", rs!DTIM_LIS_FECHA_INICIO)
           list_item.SubItems(3) = IIf(IsNull(rs!DTIM_LIS_FECHA_FIN), "", rs!DTIM_LIS_FECHA_FIN)
           list_item.SubItems(4) = IIf(IsNull(rs!vcha_mon_moneda), "", rs!vcha_mon_moneda)
           rs.MoveNext:
           numero_items_listadeprecios = numero_items_listadeprecios + 1
       Wend
       rs.Close
    End If
    rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + txt_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       txt_nombre_moneda = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
    Else
       txt_nombre_moneda = ""
    End If
    rs.Close
End Sub


Sub pro_textos()
'On Error GoTo err0:
Dim var_n As Double
   var_n = lv_listadeprecios.ListItems.Count
   If var_n > 0 Then
      txt_lista_precios = lv_listadeprecios.selectedItem
      txt_nombre_lista_precios = lv_listadeprecios.selectedItem.SubItems(1)
      txt_fecha_inicio = lv_listadeprecios.selectedItem.SubItems(2)
      txt_fecha_fin = lv_listadeprecios.selectedItem.SubItems(3)
      txt_moneda = lv_listadeprecios.selectedItem.SubItems(4)
      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + txt_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_moneda = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
      Else
         txt_nombre_moneda = ""
      End If
      rs.Close
   End If
   var_numero_renglones = lv_listadeprecios.Height / 312.5
   If var_n > var_numero_renglones Then
      lv_listadeprecios.ColumnHeaders(2).Width = 3850
   Else
      lv_listadeprecios.ColumnHeaders(2).Width = 4099.71
   End If
   var_modifica_registro_lista_precios = True
   var_hubo_cambios = False
   Me.txt_lista_precios.Enabled = False
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_lista_precios = False Then
        Set list_item = lv_listadeprecios.ListItems.Add(, , txt_lista_precios)
        list_item.SubItems(1) = txt_nombre_lista_precios
        list_item.SubItems(2) = txt_fecha_inicio
        list_item.SubItems(3) = txt_fecha_fin
        list_item.SubItems(4) = txt_moneda
        list_item.EnsureVisible
        list_item.Selected = True
        numero_items_listadeprecios = numero_items_listadeprecios + 1
    Else
        lv_listadeprecios.ListItems.Item(lv_listadeprecios.selectedItem.Index).Checked = False
        lv_listadeprecios.ListItems.Item(lv_listadeprecios.selectedItem.Index) = txt_lista_precios
        lv_listadeprecios.ListItems.Item(lv_listadeprecios.selectedItem.Index).ListSubItems(1) = txt_nombre_lista_precios
        lv_listadeprecios.ListItems.Item(lv_listadeprecios.selectedItem.Index).ListSubItems(2) = txt_fecha_inicio
        lv_listadeprecios.ListItems.Item(lv_listadeprecios.selectedItem.Index).ListSubItems(3) = txt_fecha_fin
        lv_listadeprecios.ListItems.Item(lv_listadeprecios.selectedItem.Index).ListSubItems(4) = txt_moneda
        lv_listadeprecios.ListItems.Item(lv_listadeprecios.selectedItem.Index).Selected = True
    End If
    lv_listadeprecios.SetFocus
End Sub


Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_listadeprecios, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub

Private Sub txt_fecha_fin_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fecha_fin_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_fecha_inicio_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_fecha_inicio_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_lista_precios_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_lista_precios_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_moneda_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_moneda_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_moneda_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_monedas order by vcha_mon_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_mon_moneda_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MONEDAS"
      VAR_TIPO_LISTA = 1
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
   If KeyCode = 117 Then
      var_catalogo_articulos = True
      frmmonedas.Show
   End If
End Sub

Private Sub txt_moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_moneda_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
   If Trim(txt_moneda) <> "" Then
      rs.Open "select * from tb_monedas where vcha_mon_moneda_id = '" + txt_moneda + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_moneda = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
      Else
         MsgBox "Clave de moneda incorrecta", vbOKOnly, "ATENCION"
         txt_moneda = ""
         txt_nombre_moneda = ""
      End If
      rs.Close
   Else
      txt_moneda = ""
      txt_nombre_moneda = ""
   End If
End Sub

Private Sub txt_nombre_lista_precios_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_lista_precios_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_nombre_moneda_Change()
   var_hubo_cambios = True
End Sub

Private Sub txt_nombre_moneda_GotFocus()
   Frmmenu2.StatusBar1.Panels(1) = "Presione F5 para ver la información disponible y F6 para accesar al catálogo"
End Sub

Private Sub txt_nombre_moneda_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_monedas order by vcha_mon_nombre", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_mon_moneda_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_mon_nombre), "", rs!vcha_mon_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "MONEDAS"
      VAR_TIPO_LISTA = 1
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
   If KeyCode = 117 Then
      var_catalogo_articulos = True
      frmmonedas.Show
   End If
End Sub

Private Sub txt_nombre_moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      If Me.cmd_guardar.Enabled = True Then
         Me.cmd_guardar.SetFocus
      End If
   End If
End Sub

Private Sub txt_nombre_moneda_LostFocus()
   Frmmenu2.StatusBar1.Panels(1) = ""
End Sub
