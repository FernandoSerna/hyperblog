VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmpaqueteria_odm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresi�n de guias de Omnibus de M�xico"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   945
      TabIndex        =   36
      Top             =   585
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   37
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
         TabIndex        =   38
         Top             =   135
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6345
      Picture         =   "frmpaqueteria_odm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmpaqueteria_odm.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Picture         =   "frmpaqueteria_odm.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   105
      TabIndex        =   19
      Top             =   270
      Width           =   6600
   End
   Begin VB.Frame Frame1 
      Height          =   5610
      Left            =   120
      TabIndex        =   20
      Top             =   390
      Width           =   6585
      Begin VB.TextBox txt_contable 
         Height          =   315
         Left            =   1275
         TabIndex        =   18
         Top             =   4875
         Width           =   1800
      End
      Begin VB.TextBox txt_nombre_cliente 
         Height          =   315
         Left            =   2835
         TabIndex        =   4
         Top             =   165
         Width           =   3645
      End
      Begin VB.TextBox txt_telefono 
         Height          =   315
         Left            =   4665
         TabIndex        =   15
         Top             =   3840
         Width           =   1800
      End
      Begin VB.TextBox txt_orden 
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
         Height          =   420
         Left            =   1275
         TabIndex        =   5
         Top             =   510
         Width           =   1575
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1275
         TabIndex        =   3
         Top             =   165
         Width           =   1545
      End
      Begin VB.TextBox txt_direccion 
         Height          =   315
         Left            =   1275
         TabIndex        =   8
         Top             =   1770
         Width           =   5190
      End
      Begin VB.TextBox txt_colonia 
         Height          =   315
         Left            =   1275
         TabIndex        =   9
         Top             =   2115
         Width           =   5190
      End
      Begin VB.TextBox txt_ciudad 
         Height          =   315
         Left            =   1275
         TabIndex        =   10
         Top             =   2460
         Width           =   5190
      End
      Begin VB.TextBox txt_estado 
         Height          =   315
         Left            =   1275
         TabIndex        =   11
         Top             =   2805
         Width           =   5190
      End
      Begin VB.TextBox txt_municipio 
         Height          =   315
         Left            =   1275
         TabIndex        =   12
         Top             =   3150
         Width           =   5190
      End
      Begin VB.TextBox txt_pais 
         Height          =   315
         Left            =   1275
         TabIndex        =   13
         Top             =   3495
         Width           =   5190
      End
      Begin VB.TextBox txt_cp 
         Height          =   315
         Left            =   1275
         TabIndex        =   14
         Top             =   3840
         Width           =   1800
      End
      Begin VB.TextBox txt_observacion 
         Height          =   315
         Left            =   1275
         TabIndex        =   16
         Top             =   4185
         Width           =   5190
      End
      Begin VB.TextBox txt_guia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1275
         TabIndex        =   6
         Top             =   960
         Width           =   3000
      End
      Begin VB.TextBox txt_paquetes 
         Height          =   315
         Left            =   1275
         TabIndex        =   17
         Top             =   4530
         Width           =   1800
      End
      Begin VB.TextBox txt_representante 
         Height          =   315
         Left            =   1275
         TabIndex        =   7
         Top             =   1425
         Width           =   5190
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   4320
         Picture         =   "frmpaqueteria_odm.frx":083E
         ScaleHeight     =   960
         ScaleWidth      =   2145
         TabIndex        =   21
         Top             =   4530
         Width           =   2145
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Contable:"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   4935
         Width           =   675
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tel�fono:"
         Height          =   195
         Left            =   3885
         TabIndex        =   35
         Top             =   3900
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden:"
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   150
         TabIndex        =   33
         Top             =   225
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Calle y n�mero:"
         Height          =   195
         Left            =   150
         TabIndex        =   32
         Top             =   1830
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   2175
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   2865
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   3210
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pais:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   3555
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "C.P.:"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   3900
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Observaci�n:"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   4245
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Guia:"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Paquetes:"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   4590
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Representante:"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   1485
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmpaqueteria_odm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_TIPO_LISTA As Integer
Private Sub cmd_imprimir_Click()
   If Me.txt_cliente <> "" Then
      var_si = MsgBox("Se va a imprimir la guia", vbYesNo, "ATENCION")
      If var_si = 6 Then
         'If IsNumeric(Me.txt_orden) Then
            Open (App.Path & "\MUPA_GUIA_" + Trim(Me.txt_guia) + ".bat") For Output As #2
            Open (App.Path & "\MUPA_GUIA_" + Trim(Me.txt_guia) + ".txt") For Output As #1
            Print #1, Chr(15) + Chr(27) + Chr(64)
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, Spc(74); Date
            Print #1, ""
            Print #1, Spc(50); Me.txt_ciudad
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, Spc(60); Me.txt_representante
            Print #1, ""
            Print #1, Spc(60); Me.txt_direccion
            Print #1, ""
            Print #1, Spc(60); Me.txt_colonia
            Print #1, Spc(100); Me.txt_cp
            var_municipio = Mid(Me.txt_municipio, 1, 30)
            If Len(var_municipio) < 30 Then
               For j_n = Len(var_municipio) To 30
                   var_municipio = var_municipio + " "
               Next j_n
            End If
            var_municipio = var_municipio + "    " + Me.txt_telefono
            Print #1, Spc(60); var_municipio
      
            Print #1, ""
            Print #1, Spc(30); "BLANCOS"
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, Spc(10); Me.txt_contable
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Close #1
            Print #2, "copy " + App.Path + "\MUPA_GUIA_" + Trim(Me.txt_guia) + ".txt lpt1"
            Close #2
            var_Archivo = App.Path & "\MUPA_GUIA_" + Trim(Me.txt_guia) + ".bat"
            x = Shell(var_Archivo, vbHide)
         'End If
      End If
   Else
      MsgBox "No se a seleccionado un cliente", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_orden = ""
   Me.txt_guia = ""
   Me.txt_cliente = ""
   Me.txt_nombre_cliente = ""
   Me.txt_representante = ""
   Me.txt_colonia = ""
   Me.txt_direccion = ""
   Me.txt_ciudad = ""
   Me.txt_estado = ""
   Me.txt_municipio = ""
   Me.txt_municipio = ""
   Me.txt_pais = ""
   Me.txt_cp = ""
   Me.txt_observacion = ""
   Me.txt_orden = ""
   Me.txt_telefono = ""
   Me.txt_contable = ""
   Me.txt_cliente.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 73 Then
      cmd_imprimir_Click
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If VAR_TIPO_LISTA = 0 Then
         Unload Me
      Else
         Me.frm_lista.Visible = False
         Me.txt_cliente.SetFocus
      End If
   End If
End Sub

Private Sub Form_Load()
   VAR_TIPO_LISTA = 0
   Top = 1000
   Left = 2500
   Me.frm_lista.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_packing_list)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_cliente = lv_lista.selectedItem
      Me.txt_nombre_cliente = lv_lista.selectedItem.SubItems(1)
      Me.txt_cliente.SetFocus
   End If
End Sub

Private Sub lv_lista_LostFocus()
   VAR_TIPO_LISTA = 0
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_cliente_Change()
   Me.txt_nombre_cliente = ""
   Me.txt_representante = ""
   Me.txt_direccion = ""
   Me.txt_colonia = ""
   Me.txt_estado = ""
   Me.txt_ciudad = ""
   Me.txt_pais = ""
   Me.txt_municipio = ""
   Me.txt_cp = ""
   Me.txt_telefono = ""
   Me.txt_orden = ""
   Me.txt_guia = ""
   Me.txt_observacion = ""
   Me.txt_contable = ""
End Sub

Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If VAR_UNIDAD_ORGANIZACIONAL_ID = "23" Then
         rs.Open "select * from TB_CLIENTES WHERE (VCHA_TCL_TIPO_CLIENTE_ID = 'FT') ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "select * from TB_CLIENTES WHERE (VCHA_TCL_TIPO_CLIENTE_ID = 'T') AND (VCHA_TIT_TITULAR_ID = 'T000000444') ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
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
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cliente_LostFocus()
   If Me.txt_cliente <> "" Then
      rsaux.Open "select * from vw_clientes where vcha_cli_clave_id = '" + Me.txt_cliente + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux.EOF Then
         Me.txt_nombre_cliente = IIf(IsNull(rsaux!VCHA_CLI_NOMBRE), "", rsaux!VCHA_CLI_NOMBRE)
         Me.txt_representante = IIf(IsNull(rsaux!vcha_cli_representante), "", rsaux!vcha_cli_representante)
         Me.txt_direccion = IIf(IsNull(rsaux!vcha_cli_direccion), "", rsaux!vcha_cli_direccion)
         Me.txt_colonia = IIf(IsNull(rsaux!vcha_col_nombre), "", rsaux!vcha_col_nombre)
         Me.txt_estado = IIf(IsNull(rsaux!vcha_est_nombre), "", rsaux!vcha_est_nombre)
         Me.txt_ciudad = IIf(IsNull(rsaux!vcha_ciu_nombre), "", rsaux!vcha_ciu_nombre)
         Me.txt_pais = IIf(IsNull(rsaux!vcha_pai_nombre), "", rsaux!vcha_pai_nombre)
         Me.txt_municipio = IIf(IsNull(rsaux!vcha_mun_nombre), "", rsaux!vcha_mun_nombre)
         Me.txt_cp = IIf(IsNull(rsaux!vcha_cli_cp), "", rsaux!vcha_cli_cp)
         Me.txt_telefono = IIf(IsNull(rsaux!vcha_cli_telefono), "", rsaux!vcha_cli_telefono)
         Me.txt_orden = ""
         Me.txt_guia = ""
         Me.txt_observacion = ""
         Me.txt_contable = IIf(IsNull(rsaux!vcha_age_contable), "", rsaux!vcha_age_contable)
      Else
         Me.txt_nombre_cliente = ""
         Me.txt_representante = ""
         Me.txt_direccion = ""
         Me.txt_colonia = ""
         Me.txt_estado = ""
         Me.txt_ciudad = ""
         Me.txt_pais = ""
         Me.txt_municipio = ""
         Me.txt_cp = ""
         Me.txt_telefono = ""
         Me.txt_orden = ""
         Me.txt_guia = ""
         Me.txt_observacion = ""
         Me.txt_contable = ""
         MsgBox "Clave de cliente incorrecta", vbOKOnly, "ATENCION"
      End If
      rsaux.Close
   Else
      Me.txt_nombre_cliente = ""
      Me.txt_representante = ""
      Me.txt_direccion = ""
      Me.txt_colonia = ""
      Me.txt_estado = ""
      Me.txt_ciudad = ""
      Me.txt_pais = ""
      Me.txt_municipio = ""
      Me.txt_cp = ""
      Me.txt_telefono = ""
      Me.txt_orden = ""
      Me.txt_guia = ""
      Me.txt_observacion = ""
      Me.txt_contable = ""
   End If
End Sub

Private Sub txt_nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      lv_lista.ListItems.Clear
      If VAR_UNIDAD_ORGANIZACIONAL_ID = "23" Then
         rs.Open "select * from TB_CLIENTES WHERE (VCHA_TCL_TIPO_CLIENTE_ID = 'FT') ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      Else
         rs.Open "select * from TB_CLIENTES WHERE (VCHA_TCL_TIPO_CLIENTE_ID = 'T') AND (VCHA_TIT_TITULAR_ID = 'T000000444') ORDER BY VCHA_CLI_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
      End If
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cli_clave_id)
            list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CLIENTES"
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
End Sub

Private Sub txt_nombre_cliente_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_orden_Change()
   Me.txt_guia = ""
   Me.txt_observacion = ""
   Me.txt_paquetes = ""
End Sub

Private Sub txt_orden_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_orden) Then
         rs.Open "select * from tb_enc_orden_surtido where inte_ors_orden_surtido = " + Me.txt_orden, cnn, adOpenDynamic, adLockOptimistic
         If Me.txt_cliente <> "" Then
            If Me.txt_cliente = rs!vcha_cli_clave_id Then
               var_paqueteria = IIf(IsNull(rs!vcha_paq_clave_id), "", rs!vcha_paq_clave_id)
               If var_paqueteria = "020" Then
                  Me.txt_guia = IIf(IsNull(rs!vcha_paq_guia), "", rs!vcha_paq_guia)
                  rsaux.Open "SELECT VCHA_CAJ_NOMBRE, COUNT(*) AS NUMERO_CAJAS From VW_NUMERO_CAJAS_PAQUETRIA Where (INTE_ORS_ORDEN_SURTIDO = " + Me.txt_orden + ") GROUP BY VCHA_CAJ_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_cadena = ""
                     VAR_NUMERO_CAJAS = 0
                     While Not rsaux.EOF
                           If var_cadena = "" Then
                              var_cadena = var_cadena + CStr(IIf(IsNull(NUMERO_CAJAS), "", rsaux!NUMERO_CAJAS)) + " " + IIf(IsNull(rsaux!VCHA_CAJ_NOMBRE), "", rsaux!VCHA_CAJ_NOMBRE)
                           Else
                              var_cadena = var_cadena + ", " + CStr(IIf(IsNull(NUMERO_CAJAS), "", rsaux!NUMERO_CAJAS)) + " " + IIf(IsNull(rsaux!VCHA_CAJ_NOMBRE), "", rsaux!VCHA_CAJ_NOMBRE)
                           End If
                           VAR_NUMERO_CAJAS = VAR_NUMERO_CAJAS + 1
                           rsaux.MoveNext
                     Wend
                     Me.txt_observacion = var_cadena
                     Me.txt_paquetes = VAR_NUMERO_CAJAS
                  Else
                     MsgBox "La mercanc�a no a sido empacada", vbOKOnly, "ATENCION"
                  End If
                  rsaux.Close
               Else
                  MsgBox "La orden de surtido sera enviada por " + IIf(IsNull(rs!vcha_paq_nombre), "", rs!vcha_paq_nombre), vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "El cliente seleccionado no corresponde al cliente de la orden de surtido", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Debe de seleccionar un cliente", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         MsgBox "N�mero de orden surtido incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

