VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_rutas_distribucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rutas"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton cmd_migrar_oracle 
      Caption         =   "ORACLE"
      Height          =   315
      Left            =   690
      TabIndex        =   18
      Top             =   15
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9285
      Picture         =   "frmoracle_rutas_distribucion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   30
      TabIndex        =   11
      Top             =   345
      Width           =   9600
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   30
      Picture         =   "frmoracle_rutas_distribucion.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   3720
      Left            =   30
      TabIndex        =   9
      Top             =   2970
      Width           =   9630
      Begin MSComctlLib.ListView lv_rutas 
         Height          =   3135
         Left            =   60
         TabIndex        =   0
         Top             =   495
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5530
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
         NumItems        =   25
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre Ruta"
            Object.Width           =   11465
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Lunes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Martes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Miercoles"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Jueves"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Viernes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Sábado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Entrega Lunes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Entrega Martes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Entrega Miercoles"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Entrega Jueves"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Entrega Viernes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Entrega Sabado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Carga Lunes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Carga Martes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Carga Miercoles"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Carga Jueves"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Carga Viernes"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Carga sabado"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Paqueteria"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "domingo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "carga domingo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "entrega domingo"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   " Rutas"
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   45
         TabIndex        =   10
         Top             =   135
         Width           =   9525
      End
   End
   Begin VB.CommandButton com_guardar 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   360
      Picture         =   "frmoracle_rutas_distribucion.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   30
      TabIndex        =   6
      Top             =   405
      Width           =   9615
      Begin VB.CheckBox chk_domingo 
         Caption         =   "Domingo"
         Height          =   405
         Left            =   8280
         TabIndex        =   37
         Top             =   1080
         Width           =   1080
      End
      Begin VB.CheckBox chk_entrega_domingo 
         Caption         =   "Domingo"
         Height          =   405
         Left            =   8295
         TabIndex        =   36
         Top             =   1965
         Width           =   1080
      End
      Begin VB.CheckBox chk_carga_domingo 
         Caption         =   "Domingo"
         Height          =   405
         Left            =   8295
         TabIndex        =   35
         Top             =   1575
         Width           =   1080
      End
      Begin VB.CheckBox chk_paqueteria 
         Caption         =   "Paqueteria"
         Height          =   285
         Left            =   2850
         TabIndex        =   34
         Top             =   210
         Width           =   1155
      End
      Begin VB.CheckBox chk_carga_sabado 
         Caption         =   "Sabado"
         Height          =   405
         Left            =   7140
         TabIndex        =   32
         Top             =   1575
         Width           =   1080
      End
      Begin VB.CheckBox chk_carga_viernes 
         Caption         =   "Viernes"
         Height          =   405
         Left            =   5850
         TabIndex        =   31
         Top             =   1575
         Width           =   990
      End
      Begin VB.CheckBox chk_carga_jueves 
         Caption         =   "Jueves"
         Height          =   405
         Left            =   4620
         TabIndex        =   30
         Top             =   1575
         Width           =   990
      End
      Begin VB.CheckBox chk_carga_miercoles 
         Caption         =   "Miercoles"
         Height          =   405
         Left            =   3300
         TabIndex        =   29
         Top             =   1575
         Width           =   990
      End
      Begin VB.CheckBox chk_carga_martes 
         Caption         =   "Martes"
         Height          =   405
         Left            =   2220
         TabIndex        =   28
         Top             =   1575
         Width           =   990
      End
      Begin VB.CheckBox chk_carga_lunes 
         Caption         =   "Lunes"
         Height          =   405
         Left            =   1170
         TabIndex        =   27
         Top             =   1575
         Width           =   990
      End
      Begin VB.CheckBox chk_entrega_sabado 
         Caption         =   "Sabado"
         Height          =   405
         Left            =   7140
         TabIndex        =   26
         Top             =   1965
         Width           =   1080
      End
      Begin VB.CheckBox chk_entrega_viernes 
         Caption         =   "Viernes"
         Height          =   405
         Left            =   5850
         TabIndex        =   25
         Top             =   1965
         Width           =   990
      End
      Begin VB.CheckBox chk_entrega_jueves 
         Caption         =   "Jueves"
         Height          =   405
         Left            =   4620
         TabIndex        =   24
         Top             =   1965
         Width           =   990
      End
      Begin VB.CheckBox chk_entrega_miercoles 
         Caption         =   "Miercoles"
         Height          =   405
         Left            =   3300
         TabIndex        =   23
         Top             =   1965
         Width           =   990
      End
      Begin VB.CheckBox chk_entrega_martes 
         Caption         =   "Martes"
         Height          =   405
         Left            =   2190
         TabIndex        =   22
         Top             =   1965
         Width           =   990
      End
      Begin VB.CheckBox chk_entrega_lunes 
         Caption         =   "Lunes"
         Height          =   405
         Left            =   1170
         TabIndex        =   21
         Top             =   1965
         Width           =   990
      End
      Begin VB.CheckBox chk_sabado 
         Caption         =   "Sabado"
         Height          =   405
         Left            =   7125
         TabIndex        =   17
         Top             =   1080
         Width           =   1080
      End
      Begin VB.CheckBox chk_viernes 
         Caption         =   "Viernes"
         Height          =   405
         Left            =   5865
         TabIndex        =   16
         Top             =   1080
         Width           =   990
      End
      Begin VB.CheckBox chk_jueves 
         Caption         =   "Jueves"
         Height          =   405
         Left            =   4620
         TabIndex        =   15
         Top             =   1080
         Width           =   990
      End
      Begin VB.CheckBox chk_miercoles 
         Caption         =   "Miercoles"
         Height          =   405
         Left            =   3300
         TabIndex        =   14
         Top             =   1080
         Width           =   990
      End
      Begin VB.CheckBox chk_martes 
         Caption         =   "Martes"
         Height          =   405
         Left            =   2220
         TabIndex        =   13
         Top             =   1080
         Width           =   990
      End
      Begin VB.CheckBox chk_lunes 
         Caption         =   "Lunes"
         Height          =   405
         Left            =   1185
         TabIndex        =   12
         Top             =   1080
         Width           =   990
      End
      Begin VB.TextBox txt_clave 
         Height          =   390
         Left            =   930
         TabIndex        =   4
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox txt_nombre 
         Height          =   420
         Left            =   945
         TabIndex        =   5
         Top             =   600
         Width           =   8550
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dia pedido:"
         Height          =   195
         Left            =   150
         TabIndex        =   33
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dia entrega:"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   1965
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dia carga:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   1575
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   675
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmoracle_rutas_distribucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_dias_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.com_guardar.SetFocus
   End If
End Sub

Private Sub cmd_migrar_oracle_Click()
   rs.Open "SELECT * FROM tB_RUTAS_DISTRIBUCION", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         rsaux.Open "select * from XXVIA_TB_RUTAS_DISTRIBUCION where ruta = '" + rs!ruta + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If rsaux.EOF Then
            rsaux1.Open "insert into xxvia_Tb_rutas_distribucion (ruta, nombre_ruta, lunes, martes, miercoles, jueves, viernes, sabado, paqueteria) values ('" + rs!ruta + "','" + rs!nombre_ruta + "'," + CStr(IIf(IsNull(rs!lunes), 0, rs!lunes)) + "," + CStr(IIf(IsNull(rs!martes), 0, rs!martes)) + "," + CStr(IIf(IsNull(rs!miercoles), 0, rs!miercoles)) + "," + CStr(IIf(IsNull(rs!jueves), 0, rs!jueves)) + "," + CStr(IIf(IsNull(rs!viernes), 0, rs!viernes)) + "," + CStr(IIf(IsNull(rs!sabado), 0, rs!sabado)) + "," + CStr(chk_paqueteria) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
         Else
            rsaux1.Open "update xxvia_Tb_rutas_distribucion set nombre_ruta = '" + rs!nombre_ruta + "', lunes = " = CStr(IIf(IsNull(rs!lunes), 0, rs!lunes)) + ", martes = " + CStr(IIf(IsNull(rs!martes), 0, rs!martes)) + ", miercoles = " + CStr(IIf(IsNull(rs!miercoles), 0, rs!miercoles)) + ",jueves = " + CStr(IIf(IsNull(rs!jueves), 0, rs!jueves)) + ", viernes = " + CStr(IIf(IsNull(rs!viernes), 0, rs!viernes)) + ", paqueteria = " + CStr(chk_paqueteria) + " where ruta = '" + rs!ruta + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_clave = ""
   Me.txt_nombre = ""
   Me.txt_clave.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   
End Sub

Private Sub com_guardar_Click()
   If Me.txt_clave <> "" Then
      If Me.txt_nombre <> "" Then
         'If Me.cmb_dias.Text = "LUNES" Or Me.cmb_dias.Text = "MARTES" Or Me.cmb_dias.Text = "MIERCOLES" Or Me.cmb_dias.Text = "JUEVES" Or Me.cmb_dias.Text = "VIERNES" Or Me.cmb_dias.Text = "SABADO" Then
            rs.Open "SELECT * FROM XXVIA_VW_RUTAS_DISTRIBUCION WHERE RUTA = '" + Me.txt_clave + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               'frmautoriza_cambios_distribucion.Show 1
               var_contraseña_cambios_distribucion = "X"
               If var_contraseña_cambios_distribucion <> "" Then
            
                  var_si = MsgBox("¿Desea aplicar los cambios a la ruta?", vbYesNo, "ATENCION")
                  If var_si = 6 Then
                     var_cadena = "UPDATE XXVIA_TB_RUTAS_DISTRIBUCION SET NOMBRE_RUTA = '" + Me.txt_nombre + "', LUNES = " + CStr(Me.chk_lunes) + ", MARTES = " + CStr(Me.chk_martes) + ", MIERCOLES = " + CStr(Me.chk_miercoles) + ", JUEVES = " + CStr(Me.chk_jueves) + ", VIERNES = " + CStr(Me.chk_viernes) + ", sabado = " + CStr(Me.chk_sabado) + ", ENT_LUNES = " + CStr(Me.chk_entrega_lunes) + ", ENT_MARTES = " + CStr(Me.chk_entrega_martes) + ", ENT_MIERCOLES = " + CStr(Me.chk_entrega_miercoles) + ", ENT_JUEVES = " + CStr(Me.chk_entrega_jueves) + ", ENT_VIERNES = " + CStr(Me.chk_entrega_viernes) + ", ENT_sabado = " + CStr(Me.chk_entrega_sabado) + ", car_lunes = " + CStr(Me.chk_carga_lunes) + ",car_martes = " + CStr(Me.chk_carga_martes) + ",car_miercoles = " + CStr(Me.chk_carga_miercoles) + ", car_jueves = " + CStr(Me.chk_carga_jueves) + ",car_viernes = " + CStr(Me.chk_carga_viernes)
                     var_cadena = var_cadena + ", car_sabado = " + CStr(Me.chk_carga_sabado) + ", paqueteria = " + CStr(chk_paqueteria) + ", DOMINGO  =" + CStr(Me.chk_domingo) + ", CAR_DOMINGO = " + CStr(Me.chk_carga_domingo) + ", ENT_DOMINGO = " + CStr(Me.chk_entrega_domingo) + "  WHERE RUTA = '" + Me.txt_clave + "'"
                     'MsgBox var_cadena
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     var_cadena = "INSERT INTO TB_ORACLE_BITACORA_RUTAS ([ACCION],[RUTA],[NOMBRE_RUTA],[LUNES],[MARTES],[MIERCOLES],[JUEVES],[VIERNES],[SABADO],[DOMINGO],[ENTREGA_LUNES],[ENTREGA_MARTES],[ENTREGA_MIERCOLES],[ENTREGA_JUEVES],ENTREGA_VIERNES,[ENTREGA_SABADO],[ENTREGA_DOMINGO],[CARGA_LUNES],[CARGA_MARTES],[CARGA_MIERCOLES],[CARGA_JUEVES],[CARGA_VIERNES],[CARGA_SABADO],[CARGA_DOMINGO],[USUARIO],[MAQUINA],[FECHA])"
                     var_cadena = var_cadena + " Values ('ACTUALIZO','" + Me.txt_clave + "','" + Me.txt_nombre + "'," + CStr(Me.chk_lunes) + "," + CStr(Me.chk_martes) + "," + CStr(Me.chk_miercoles) + "," + CStr(Me.chk_jueves) + "," + CStr(Me.chk_viernes) + "," + CStr(Me.chk_sabado) + "," + CStr(Me.chk_domingo) + "," + CStr(Me.chk_entrega_lunes) + "," + CStr(Me.chk_entrega_martes) + "," + CStr(Me.chk_entrega_miercoles) + "," + CStr(Me.chk_entrega_jueves) + "," + CStr(Me.chk_entrega_viernes) + "," + CStr(Me.chk_entrega_sabado) + "," + CStr(Me.chk_entrega_domingo) + "," + CStr(Me.chk_carga_lunes) + "," + CStr(Me.chk_carga_martes) + "," + CStr(Me.chk_carga_miercoles) + "," + CStr(Me.chk_carga_jueves) + "," + CStr(Me.chk_carga_viernes) + "," + CStr(Me.chk_carga_sabado) + "," + CStr(Me.chk_carga_domingo) + ",'" + var_clave_usuario_global + "','" + fun_NombrePc + "',GETDATE())"
                     rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                     
                     If Me.chk_lunes = 0 Then
                        var_lunes = "NO"
                     Else
                        var_lunes = "SI"
                     End If
                     
                     If Me.chk_martes = 0 Then
                        var_martes = "NO"
                     Else
                        var_martes = "SI"
                     End If
                     If Me.chk_miercoles = 0 Then
                        var_miercoles = "NO"
                     Else
                        var_miercoles = "SI"
                     End If
                     If Me.chk_jueves = 0 Then
                        var_jueves = "NO"
                     Else
                        var_jueves = "SI"
                     End If
                     If Me.chk_viernes = 0 Then
                        var_viernes = "NO"
                     Else
                        var_viernes = "SI"
                     End If
                     If Me.chk_sabado = 0 Then
                        var_sabado = "NO"
                     Else
                        var_sabado = "SI"
                     End If
                     If Me.chk_domingo = 0 Then
                        var_domingo = "NO"
                     Else
                        var_domingo = "SI"
                     End If
                     
                     
                     
                     If Me.chk_entrega_lunes = 0 Then
                        var_entrega_lunes = "NO"
                     Else
                        var_entrega_lunes = "SI"
                     End If
                     
                     If Me.chk_entrega_martes = 0 Then
                        var_entrega_martes = "NO"
                     Else
                        var_entrega_martes = "SI"
                     End If
                     If Me.chk_entrega_miercoles = 0 Then
                        var_entrega_miercoles = "NO"
                     Else
                        var_entrega_miercoles = "SI"
                     End If
                     If Me.chk_entrega_jueves = 0 Then
                        var_entrega_jueves = "NO"
                     Else
                        var_entrega_jueves = "SI"
                     End If
                     If Me.chk_entrega_viernes = 0 Then
                        var_entrega_viernes = "NO"
                     Else
                        var_enrega_viernes = "SI"
                     End If
                     If Me.chk_entrega_sabado = 0 Then
                        var_entrega_sabado = "NO"
                     Else
                        var_enrega_sabado = "SI"
                     End If
                     If Me.chk_entrega_domingo = 0 Then
                        var_entrega_domingo = "NO"
                     Else
                        var_entrega_domingo = "SI"
                     End If
                     
                     
                     If Me.chk_carga_lunes = 0 Then
                        var_carga_lunes = "NO"
                     Else
                        var_carga_lunes = "SI"
                     End If
                     
                     If Me.chk_carga_martes = 0 Then
                         var_carga_martes = "NO"
                     Else
                        var_carga_martes = "SI"
                     End If
                     If Me.chk_carga_miercoles = 0 Then
                        var_carga_miercoles = "NO"
                     Else
                        var_carga_miercoles = "SI"
                     End If
                     If Me.chk_carga_jueves = 0 Then
                        var_carga_jueves = "NO"
                     Else
                        var_carga_jueves = "SI"
                     End If
                     If Me.chk_carga_viernes = 0 Then
                        var_carga_viernes = "NO"
                     Else
                        var_carga_viernes = "SI"
                     End If
                     If Me.chk_carga_sabado = 0 Then
                        var_carga_sabado = "NO"
                     Else
                        var_carga_sabado = "SI"
                     End If
                     If Me.chk_carga_domingo = 0 Then
                        var_carga_domingo = "NO"
                     Else
                        var_carga_domingo = "SI"
                     End If
                     
                     
                     var_asunto = "Se informa que se a cambiado los dias de pedido de la ruta <strong>" + Me.txt_nombre + "</strong>, quedando de la siguiente manera:<br /><br />Lunes: " + var_lunes + "<br /><br />Martes: " + var_martes + "<br /><br /> Miercoles: " + var_miercoles + "<br /><br />Jueves: " + var_jueves + "<br /><br /> Viernes:" + var_viernes
                     var_asunto = var_asunto + "<br /><br /> Sabado: " + var_sabado + "<br /><br /> Domingo: " + var_domingo + "<br /><br /> Lunes entrega: " + CStr(var_entrega_lunes) + "<br /><br /> Martes entrega: " + CStr(var_entrega_martes) + "<br /><br /> Miercoles entrega: " + CStr(var_entrega_miercoles) + "<br /><br /> Jueves entrega: " + CStr(var_entrega_jueves) + "<br /><br /> Viernes entrega: " + CStr(var_entrega_viernes) + "<br /><br /> Sabado entrega: " + CStr(var_entrega_sabado) + "<br /><br /> Domingo entrega: " + CStr(var_entrega_domingo)
                     var_asunto = var_asunto + "<br /><br /> Lunes carga: " + CStr(var_carga_lunes) + "<br /><br /> Martes carga: " + CStr(var_carga_martes) + "<br /><br /> Miercoles carga: " + CStr(var_carga_miercoles) + "<br /><br /> Jueves carga: " + CStr(var_carga_jueves) + "<br /><br /> Viernes carga: " + CStr(var_carga_viernes) + "<br /><br /> Sabado carga: " + CStr(var_carga_sabado) + "<br /><br /> Domingo carga: " + CStr(var_carga_domingo) + ""
                     Me.Text1.Text = var_asunto
                     
                     var_cadena = "call xxvia_pk_correo.sp_enviar_email('','fserna@vianney.com.mx','','','Cambio de dias de pedido de la ruta " + Me.txt_nombre + "','" + var_asunto + "','')"
                     
                     rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                     VAR_Z = 1
                     For var_j = 1 To Me.lv_rutas.ListItems.Count
                         Me.lv_rutas.ListItems.Item(var_j).Selected = True
                         If Me.lv_rutas.selectedItem = Me.txt_clave Then
                            Me.lv_rutas.selectedItem.SubItems(1) = Me.txt_nombre
                            Me.lv_rutas.selectedItem.SubItems(2) = ""
                            Me.lv_rutas.selectedItem.SubItems(3) = Me.chk_lunes
                            Me.lv_rutas.selectedItem.SubItems(4) = Me.chk_martes
                            Me.lv_rutas.selectedItem.SubItems(5) = Me.chk_miercoles
                            Me.lv_rutas.selectedItem.SubItems(6) = Me.chk_jueves
                            Me.lv_rutas.selectedItem.SubItems(7) = Me.chk_viernes
                            Me.lv_rutas.selectedItem.SubItems(8) = Me.chk_sabado
                            Me.lv_rutas.selectedItem.SubItems(9) = Me.chk_entrega_lunes
                            Me.lv_rutas.selectedItem.SubItems(10) = Me.chk_entrega_martes
                            Me.lv_rutas.selectedItem.SubItems(11) = Me.chk_entrega_miercoles
                            Me.lv_rutas.selectedItem.SubItems(12) = Me.chk_entrega_jueves
                            Me.lv_rutas.selectedItem.SubItems(13) = Me.chk_entrega_viernes
                            Me.lv_rutas.selectedItem.SubItems(14) = Me.chk_entrega_sabado
                            Me.lv_rutas.selectedItem.SubItems(15) = Me.chk_carga_lunes
                            Me.lv_rutas.selectedItem.SubItems(16) = Me.chk_carga_martes
                            Me.lv_rutas.selectedItem.SubItems(17) = Me.chk_carga_miercoles
                            Me.lv_rutas.selectedItem.SubItems(18) = Me.chk_carga_jueves
                            Me.lv_rutas.selectedItem.SubItems(19) = Me.chk_carga_viernes
                            Me.lv_rutas.selectedItem.SubItems(20) = Me.chk_carga_sabado
                            Me.lv_rutas.selectedItem.SubItems(21) = Me.chk_paqueteria
                            Me.lv_rutas.selectedItem.SubItems(22) = Me.chk_domingo
                            Me.lv_rutas.selectedItem.SubItems(23) = Me.chk_carga_domingo
                            Me.lv_rutas.selectedItem.SubItems(24) = Me.chk_entrega_domingo
                            VAR_Z = var_j
                         End If
                     Next var_j
                     Me.lv_rutas.ListItems.Item(VAR_Z).EnsureVisible
                     Me.lv_rutas.ListItems.Item(VAR_Z).Selected = True
                  End If
               End If
            Else
               'MsgBox "insert into XXVIA_TB_RUTAS_DISTRIBUCION (ruta, nombre_ruta, dia, lunes, martes, miercoles, jueves, viernes, sabado, dia, ENT_lunes, ENT_martes, ENT_miercoles, ENT_jueves, ENT_viernes, ENT_sabado) values ('" + Me.txt_clave + "','" + Me.txt_nombre + "','" + "" + "'," + CStr(Me.chk_lunes) + "," + CStr(Me.chk_martes) + "," + CStr(Me.chk_miercoles) + "," + CStr(Me.chk_jueves) + "," + CStr(Me.chk_viernes) + "," + CStr(Me.chk_sabado) + "," + CStr(Me.chk_entrega_lunes) + "," + CStr(Me.chk_entrega_martes) + "," + CStr(Me.chk_entrega_miercoles) + "," + CStr(Me.chk_entrega_jueves) + "," + CStr(Me.chk_entrega_viernes) + "," + CStr(Me.chk_entrega_sabado) + ")"
               rsaux.Open "insert into XXVIA_TB_RUTAS_DISTRIBUCION (ruta, nombre_ruta, lunes, martes, miercoles, jueves, viernes, sabado, ENT_lunes, ENT_martes, ENT_miercoles, ENT_jueves, ENT_viernes, ENT_sabado, car_lunes, car_martes, car_miercoles, car_jueves, car_viernes, car_sabado) values ('" + Me.txt_clave + "','" + Me.txt_nombre + "'," + CStr(Me.chk_lunes) + "," + CStr(Me.chk_martes) + "," + CStr(Me.chk_miercoles) + "," + CStr(Me.chk_jueves) + "," + CStr(Me.chk_viernes) + "," + CStr(Me.chk_sabado) + "," + CStr(Me.chk_entrega_lunes) + "," + CStr(Me.chk_entrega_martes) + "," + CStr(Me.chk_entrega_miercoles) + "," + CStr(Me.chk_entrega_jueves) + "," + CStr(Me.chk_entrega_viernes) + "," + CStr(Me.chk_entrega_sabado) + "," + CStr(Me.chk_carga_lunes) + "," + CStr(Me.chk_carga_martes) + "," + CStr(Me.chk_carga_miercoles) + "," + CStr(Me.chk_carga_jueves) + "," + CStr(Me.chk_carga_viernes) + "," + CStr(Me.chk_carga_sabado) + ")", cnnoracle_4, adOpenDynamic, adLockOptimistic
               var_cadena = "INSERT INTO TB_ORACLE_BITACORA_RUTAS ([ACCION],[RUTA],[NOMBRE_RUTA],[LUNES],[MARTES],[MIERCOLES],[JUEVES],[VIERNES],[SABADO],[DOMINGO],[ENTREGA_LUNES],[ENTREGA_MARTES],[ENTREGA_MIERCOLES],[ENTREGA_JUEVES],ENTREGA_VIERNES,[ENTREGA_SABADO],[ENTREGA_DOMINGO],[CARGA_LUNES],[CARGA_MARTES],[CARGA_MIERCOLES],[CARGA_JUEVES],[CARGA_VIERNES],[CARGA_SABADO],[CARGA_DOMINGO],[USUARIO],[MAQUINA],[FECHA])"
               var_cadena = var_cadena + " Values ('INSERTO','" + Me.txt_clave + "','" + Me.txt_nombre + "'," + CStr(Me.chk_lunes) + "," + CStr(Me.chk_martes) + "," + CStr(Me.chk_miercoles) + "," + CStr(Me.chk_jueves) + "," + CStr(Me.chk_viernes) + "," + CStr(Me.chk_sabado) + "," + CStr(Me.chk_domingo) + "," + CStr(Me.chk_entrega_lunes) + "," + CStr(Me.chk_entrega_martes) + "," + CStr(Me.chk_entrega_miercoles) + "," + CStr(Me.chk_entrega_jueves) + "," + CStr(Me.chk_entrega_viernes) + "," + CStr(Me.chk_entrega_sabado) + "," + CStr(Me.chk_entrega_domingo) + "," + CStr(Me.chk_carga_lunes) + "," + CStr(Me.chk_carga_martes) + "," + CStr(Me.chk_carga_miercoles) + "," + CStr(Me.chk_carga_jueves) + "," + CStr(Me.chk_carga_viernes) + "," + CStr(Me.chk_carga_sabado) + "," + CStr(Me.chk_carga_domingo) + ",'" + var_clave_usuario_global + "','" + fun_NombrePc + "',GETDATE())"
               
               Set list_item = lv_rutas.ListItems.Add(, , Me.txt_clave)
               list_item.SubItems(1) = Me.txt_nombre
               list_item.SubItems(2) = ""
               Me.lv_rutas.selectedItem.SubItems(3) = Me.chk_lunes
               Me.lv_rutas.selectedItem.SubItems(4) = Me.chk_martes
               Me.lv_rutas.selectedItem.SubItems(5) = Me.chk_miercoles
               Me.lv_rutas.selectedItem.SubItems(6) = Me.chk_jueves
               Me.lv_rutas.selectedItem.SubItems(7) = Me.chk_viernes
               Me.lv_rutas.selectedItem.SubItems(8) = Me.chk_sabado
               Me.lv_rutas.selectedItem.SubItems(9) = Me.chk_entrega_lunes
               Me.lv_rutas.selectedItem.SubItems(10) = Me.chk_entrega_martes
               Me.lv_rutas.selectedItem.SubItems(11) = Me.chk_entrega_miercoles
               Me.lv_rutas.selectedItem.SubItems(12) = Me.chk_entrega_jueves
               Me.lv_rutas.selectedItem.SubItems(13) = Me.chk_entrega_viernes
               Me.lv_rutas.selectedItem.SubItems(14) = Me.chk_entrega_sabado
               Me.lv_rutas.selectedItem.SubItems(15) = Me.chk_carga_lunes
               Me.lv_rutas.selectedItem.SubItems(16) = Me.chk_carga_martes
               Me.lv_rutas.selectedItem.SubItems(17) = Me.chk_carga_miercoles
               Me.lv_rutas.selectedItem.SubItems(18) = Me.chk_carga_jueves
               Me.lv_rutas.selectedItem.SubItems(19) = Me.chk_carga_viernes
               Me.lv_rutas.selectedItem.SubItems(20) = Me.chk_carga_sabado
               Me.lv_rutas.selectedItem.SubItems(21) = Me.chk_paqueteria
               Me.lv_rutas.selectedItem.SubItems(22) = Me.chk_domingo
               Me.lv_rutas.selectedItem.SubItems(23) = Me.chk_carga_domingo
               Me.lv_rutas.selectedItem.SubItems(24) = Me.chk_entrega_domingo
               list_item.EnsureVisible
               list_item.Selected = True
               MsgBox "Se a insertado el registro", vbOKOnly, "ATENCION"
            End If
            rs.Close
         'Else
         '   MsgBox "No se a indicado un dia", vbOKOnly, "ATENCION"
         'End If
      Else
         MsgBox "No se a indicado el nombre de la ruta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "No se a indicado una ruta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Activate()
   If Me.lv_rutas.ListItems.Count > 0 Then
      Me.txt_clave.SetFocus
   End If

End Sub

Private Sub Form_Load()
   Top = 200
   Left = 1000
   rs.Open "SELECT * FROM XXVIA_VW_RUTAS_DISTRIBUCION", cnnoracle_4, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_rutas.ListItems.Add(, , rs!ruta)
         list_item.SubItems(1) = IIf(IsNull(rs!nombre_ruta), 0, rs!nombre_ruta)
         list_item.SubItems(2) = ""
         list_item.SubItems(3) = IIf(IsNull(rs!lunes), "0", rs!lunes)
         list_item.SubItems(4) = IIf(IsNull(rs!martes), "0", rs!martes)
         list_item.SubItems(5) = IIf(IsNull(rs!miercoles), "0", rs!miercoles)
         list_item.SubItems(6) = IIf(IsNull(rs!jueves), "0", rs!jueves)
         list_item.SubItems(7) = IIf(IsNull(rs!viernes), "0", rs!viernes)
         list_item.SubItems(8) = IIf(IsNull(rs!sabado), "", rs!sabado)
         list_item.SubItems(9) = IIf(IsNull(rs!ent_lunes), "0", rs!ent_lunes)
         list_item.SubItems(10) = IIf(IsNull(rs!ent_martes), "0", rs!ent_martes)
         list_item.SubItems(11) = IIf(IsNull(rs!ent_miercoles), "0", rs!ent_miercoles)
         list_item.SubItems(12) = IIf(IsNull(rs!ent_jueves), "0", rs!ent_jueves)
         list_item.SubItems(13) = IIf(IsNull(rs!ent_viernes), "0", rs!ent_viernes)
         list_item.SubItems(14) = IIf(IsNull(rs!ent_sabado), "", rs!ent_sabado)
         list_item.SubItems(15) = IIf(IsNull(rs!CAR_lunes), "0", rs!CAR_lunes)
         list_item.SubItems(16) = IIf(IsNull(rs!CAR_martes), "0", rs!CAR_martes)
         list_item.SubItems(17) = IIf(IsNull(rs!CAR_miercoles), "0", rs!CAR_miercoles)
         list_item.SubItems(18) = IIf(IsNull(rs!CAR_jueves), "0", rs!CAR_jueves)
         list_item.SubItems(19) = IIf(IsNull(rs!CAR_viernes), "0", rs!CAR_viernes)
         list_item.SubItems(20) = IIf(IsNull(rs!CAR_sabado), "", rs!CAR_sabado)
         list_item.SubItems(21) = IIf(IsNull(rs!paqueteria), "", rs!paqueteria)
         list_item.SubItems(22) = IIf(IsNull(rs!domingo), "", rs!domingo)
         list_item.SubItems(23) = IIf(IsNull(rs!CAR_DOMINGO), "", rs!CAR_DOMINGO)
         list_item.SubItems(24) = IIf(IsNull(rs!ENT_DOMINGO), "", rs!ENT_DOMINGO)
         rs.MoveNext
   Wend
   rs.Close
   x = 0
   If x = 1 Then
   If var_clave_usuario_global = "U0000000528" Then
      Me.chk_entrega_lunes.Enabled = False
      Me.chk_entrega_martes.Enabled = False
      Me.chk_entrega_miercoles.Enabled = False
      Me.chk_entrega_jueves.Enabled = False
      Me.chk_entrega_viernes.Enabled = False
      Me.chk_entrega_sabado.Enabled = False
      Me.chk_entrega_domingo.Enabled = False
   Else
      If var_clave_usuario_global <> "8" Then
         Me.chk_carga_lunes.Enabled = False
         Me.chk_carga_martes.Enabled = False
         Me.chk_carga_miercoles.Enabled = False
         Me.chk_carga_jueves.Enabled = False
         Me.chk_carga_viernes.Enabled = False
         Me.chk_carga_sabado.Enabled = False
         Me.chk_carga_domingo.Enabled = False
         Me.chk_lunes.Enabled = False
         Me.chk_martes.Enabled = False
         Me.chk_miercoles.Enabled = False
         Me.chk_jueves.Enabled = False
         Me.chk_viernes.Enabled = False
         Me.chk_sabado.Enabled = False
         Me.chk_domingo.Enabled = False
      Else
      End If
   End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub
Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
End Sub

Private Sub lv_rutas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_rutas, ColumnHeader)
End Sub

Private Sub lv_rutas_GotFocus()
   If Me.lv_rutas.ListItems.Count > 0 Then
      Me.txt_clave = Me.lv_rutas.selectedItem
      Me.txt_nombre = Me.lv_rutas.selectedItem.SubItems(1)
      var_dia = Me.lv_rutas.selectedItem.SubItems(2)
   End If
End Sub

Private Sub lv_rutas_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If Me.lv_rutas.ListItems.Count > 0 Then
      If Me.lv_rutas.selectedItem.SubItems(3) = "" Or Me.lv_rutas.selectedItem.SubItems(3) = "0" Then
         var_lunes = 0
      Else
         var_lunes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(4) = "" Or Me.lv_rutas.selectedItem.SubItems(4) = "0" Then
         var_martes = 0
      Else
         var_martes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(5) = "" Or Me.lv_rutas.selectedItem.SubItems(5) = "0" Then
         var_miercoles = 0
      Else
         var_miercoles = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(6) = "" Or Me.lv_rutas.selectedItem.SubItems(6) = "0" Then
         var_jueves = 0
      Else
         var_jueves = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(7) = "" Or Me.lv_rutas.selectedItem.SubItems(7) = "0" Then
         var_viernes = 0
      Else
         var_viernes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(8) = "" Or Me.lv_rutas.selectedItem.SubItems(8) = "0" Then
         var_sabado = 0
      Else
         var_sabado = 1
      End If
      
      If Me.lv_rutas.selectedItem.SubItems(9) = "" Or Me.lv_rutas.selectedItem.SubItems(9) = "0" Then
         var_ENT_lunes = 0
      Else
         var_ENT_lunes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(10) = "" Or Me.lv_rutas.selectedItem.SubItems(10) = "0" Then
         var_ENT_martes = 0
      Else
         var_ENT_martes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(11) = "" Or Me.lv_rutas.selectedItem.SubItems(11) = "0" Then
         var_ENT_miercoles = 0
      Else
         var_ENT_miercoles = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(12) = "" Or Me.lv_rutas.selectedItem.SubItems(12) = "0" Then
         var_ENT_jueves = 0
      Else
         var_ENT_jueves = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(13) = "" Or Me.lv_rutas.selectedItem.SubItems(13) = "0" Then
         var_ENT_viernes = 0
      Else
         var_ENT_viernes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(14) = "" Or Me.lv_rutas.selectedItem.SubItems(14) = "0" Then
         var_ENT_sabado = 0
      Else
         var_ENT_sabado = 1
      End If
      
      
      If Me.lv_rutas.selectedItem.SubItems(15) = "" Or Me.lv_rutas.selectedItem.SubItems(15) = "0" Then
         var_car_lunes = 0
      Else
         var_car_lunes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(16) = "" Or Me.lv_rutas.selectedItem.SubItems(16) = "0" Then
         var_car_martes = 0
      Else
         var_car_martes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(17) = "" Or Me.lv_rutas.selectedItem.SubItems(17) = "0" Then
         var_car_miercoles = 0
      Else
         var_car_miercoles = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(18) = "" Or Me.lv_rutas.selectedItem.SubItems(18) = "0" Then
         var_car_jueves = 0
      Else
         var_car_jueves = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(19) = "" Or Me.lv_rutas.selectedItem.SubItems(19) = "0" Then
         var_car_viernes = 0
      Else
         var_car_viernes = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(20) = "" Or Me.lv_rutas.selectedItem.SubItems(20) = "0" Then
         var_car_sabado = 0
      Else
         var_car_sabado = 1
      End If
      
      If Me.lv_rutas.selectedItem.SubItems(21) = "" Or Me.lv_rutas.selectedItem.SubItems(21) = "0" Then
         var_paqueteria = 0
      Else
         var_paqueteria = 1
      End If
      
      If Me.lv_rutas.selectedItem.SubItems(22) = "" Or Me.lv_rutas.selectedItem.SubItems(22) = "0" Then
         var_domingo = 0
      Else
         var_domingo = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(23) = "" Or Me.lv_rutas.selectedItem.SubItems(23) = "0" Then
         var_car_domingo = 0
      Else
         var_car_domingo = 1
      End If
      If Me.lv_rutas.selectedItem.SubItems(24) = "" Or Me.lv_rutas.selectedItem.SubItems(24) = "0" Then
         var_ent_domingo = 0
      Else
         var_ent_domingo = 1
      End If
      
      
      
      
      Me.txt_clave = Me.lv_rutas.selectedItem
      Me.txt_nombre = Me.lv_rutas.selectedItem.SubItems(1)
      'Me.cmb_dias = Me.lv_rutas.selectedItem.SubItems(2)
      Me.chk_lunes = var_lunes
      Me.chk_martes = var_martes
      Me.chk_miercoles = var_miercoles
      Me.chk_jueves = var_jueves
      Me.chk_viernes = var_viernes
      Me.chk_sabado = var_sabado
      Me.chk_entrega_lunes = var_ENT_lunes
      Me.chk_entrega_martes = var_ENT_martes
      Me.chk_entrega_miercoles = var_ENT_miercoles
      Me.chk_entrega_jueves = var_ENT_jueves
      Me.chk_entrega_viernes = var_ENT_viernes
      Me.chk_entrega_sabado = var_ENT_sabado
      Me.chk_carga_lunes = var_car_lunes
      Me.chk_carga_martes = var_car_martes
      Me.chk_carga_miercoles = var_car_miercoles
      Me.chk_carga_jueves = var_car_jueves
      Me.chk_carga_viernes = var_car_viernes
      Me.chk_carga_sabado = var_car_sabado
      Me.chk_paqueteria = var_paqueteria
      Me.chk_domingo = var_domingo
      Me.chk_carga_domingo = var_car_domingo
      Me.chk_entrega_domingo = var_ent_domingo
   End If
End Sub

Private Sub lv_rutas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ruta_distribucion = Me.lv_rutas.selectedItem
      var_nombre_ruta_distribucion = Me.lv_rutas.selectedItem.SubItems(1)
      frmoracle_asignar_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub txt_clave_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ruta_distribucion = Me.lv_rutas.selectedItem
      var_nombre_ruta_distribucion = Me.lv_rutas.selectedItem.SubItems(1)
      frmoracle_asignar_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
      Me.txt_nombre.SetFocus
   End If
End Sub

Private Sub txt_nombre_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ruta_distribucion = Me.lv_rutas.selectedItem
      var_nombre_ruta_distribucion = Me.lv_rutas.selectedItem.SubItems(1)
      frmoracle_asignar_clientes_rutas_distribucion.Show 1
   End If
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
       'Me.cmb_dias.SetFocus
   End If
End Sub
