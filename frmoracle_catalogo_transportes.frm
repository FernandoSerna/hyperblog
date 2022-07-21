VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_catalogo_transportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de transportes"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   960
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   10425
      Picture         =   "frmoracle_catalogo_transportes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   45
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmoracle_catalogo_transportes.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Guardar Alt + G"
      Top             =   60
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmoracle_catalogo_transportes.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   60
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Transportes "
      Height          =   2760
      Left            =   75
      TabIndex        =   20
      Top             =   480
      Width           =   10695
      Begin VB.TextBox txt_placa 
         Height          =   315
         Left            =   6990
         TabIndex        =   18
         Top             =   2280
         Width           =   1155
      End
      Begin VB.TextBox txt_subtiporem 
         Height          =   315
         Left            =   6990
         TabIndex        =   17
         Top             =   1920
         Width           =   3555
      End
      Begin VB.TextBox txt_aniomodelovm 
         Height          =   315
         Left            =   6990
         TabIndex        =   16
         Top             =   1560
         Width           =   1155
      End
      Begin VB.TextBox txt_placavm 
         Height          =   315
         Left            =   6990
         TabIndex        =   15
         Top             =   1200
         Width           =   3555
      End
      Begin VB.TextBox txt_configvehicular 
         Height          =   315
         Left            =   6990
         TabIndex        =   14
         Top             =   840
         Width           =   3555
      End
      Begin VB.TextBox txt_numpolizaseg 
         Height          =   315
         Left            =   6990
         TabIndex        =   13
         Top             =   480
         Width           =   3555
      End
      Begin VB.TextBox txt_nombreaseg 
         Height          =   315
         Left            =   6990
         TabIndex        =   12
         Top             =   120
         Width           =   3555
      End
      Begin VB.TextBox txt_numpermisosct 
         Height          =   315
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2265
         Width           =   1635
      End
      Begin VB.TextBox txt_permsct 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2300
         Width           =   1635
      End
      Begin VB.CheckBox chk_exportaciones 
         Caption         =   "Exportaciones"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   1590
         Width           =   2055
      End
      Begin VB.ComboBox cmb_estatus 
         Height          =   315
         ItemData        =   "frmoracle_catalogo_transportes.frx":083E
         Left            =   3720
         List            =   "frmoracle_catalogo_transportes.frx":0848
         TabIndex        =   6
         Top             =   840
         Width           =   1755
      End
      Begin VB.TextBox txt_tipo 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1920
         Width           =   4155
      End
      Begin VB.TextBox txt_rendimiento 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1560
         Width           =   1545
      End
      Begin VB.TextBox txt_placas 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1200
         Width           =   1545
      End
      Begin VB.TextBox txt_volumen 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   5
         Top             =   840
         Width           =   1545
      End
      Begin VB.TextBox txt_nombre 
         Height          =   315
         Left            =   1290
         TabIndex        =   4
         Top             =   500
         Width           =   4155
      End
      Begin VB.TextBox txt_clave 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   3
         Top             =   160
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Placa remolque:"
         Height          =   195
         Index           =   14
         Left            =   5640
         TabIndex        =   39
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "S.T. Remolque:"
         Height          =   195
         Index           =   13
         Left            =   5640
         TabIndex        =   38
         Top             =   1980
         Width           =   1110
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
         Height          =   195
         Index           =   12
         Left            =   5640
         TabIndex        =   37
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Placa VM:"
         Height          =   195
         Index           =   11
         Left            =   5640
         TabIndex        =   36
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Config. Vehicular:"
         Height          =   195
         Index           =   10
         Left            =   5640
         TabIndex        =   35
         Top             =   900
         Width           =   1245
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Póliza:"
         Height          =   195
         Index           =   9
         Left            =   5640
         TabIndex        =   34
         Top             =   540
         Width           =   465
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Aseguradora:"
         Height          =   195
         Index           =   8
         Left            =   5640
         TabIndex        =   33
         Top             =   180
         Width           =   945
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "N.P. SCT:"
         Height          =   195
         Index           =   7
         Left            =   3000
         TabIndex        =   32
         Top             =   2265
         Width           =   720
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "P. SCT:"
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   31
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estatus:"
         Height          =   195
         Left            =   2940
         TabIndex        =   29
         Top             =   900
         Width           =   570
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   28
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Rendimiento:"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   27
         Top             =   1620
         Width           =   930
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Placas:"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   26
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Volumen:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   25
         Top             =   900
         Width           =   660
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   22
         Top             =   540
         Width           =   600
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   21
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3900
      Left            =   75
      TabIndex        =   1
      Top             =   3255
      Width           =   10695
      Begin MSComctlLib.ListView lv_transportes 
         Height          =   3720
         Left            =   45
         TabIndex        =   19
         Top             =   135
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   6562
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
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   16316
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Volumen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Estatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Placas"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Rendimiento"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Exportaciones"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "permsct"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "numpermisosct"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "nombreaseg"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "numpolizaseg"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "configvehicular"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "placavm"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "aniomodelovm"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "subtiporem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "placa"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   75
      TabIndex        =   23
      Top             =   300
      Visible         =   0   'False
      Width           =   10695
   End
End
Attribute VB_Name = "frmoracle_catalogo_transportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objConn As New ADODB.Connection
Dim objCmd As New ADODB.Command
Dim objParm As ADODB.Parameter
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim var_hubo_cambios As Boolean
Dim numero_items_lineas As Integer
Dim bitacora As Boolean




Private Sub cmb_estatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_placas.SetFocus
   End If
End Sub

Private Sub cmd_guardar_Click()
        Dim var_posible As Boolean
        If Me.txt_clave <> "" Then
           If Me.txt_nombre <> "" Then
              If IsNumeric(Me.txt_volumen) Then
                 If IsNumeric(Me.txt_rendimiento) Then
                    If Me.txt_placas <> "" Then
                       If Me.txt_tipo <> "" Then
                          If Me.cmb_estatus <> "" Then
                             rs.Open "select * from tb_oracle_transportes where clave = '" + Me.txt_clave + "'", cnn, adOpenDynamic, adLockOptimistic
                             If Not rs.EOF Then
                                 var_si = MsgBox("Se va a modificar el registro", vbYesNo, "ATENCION")
                                If var_si = 6 Then
                                   rsaux.Open "UPDATE tb_oracle_transportes SET nombre = '" + Me.txt_nombre + "', VOLUMEN = " + Me.txt_volumen + ",PLACAS = '" + Me.txt_placas + "', RENDIMIENTO = " + Me.txt_rendimiento + ", ESTATUS = '" + Me.cmb_estatus.Text + "', TIPO = '" + Me.txt_tipo + "', exportaciones = " + CStr(Me.chk_exportaciones) + " WHERE clave = '" + Trim(Me.txt_clave) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   var_modifica_registro_linea = False
                                   Call pro_actualiza_ListView
                                End If
                             Else
                                rsaux.Open "INSERT INTO tb_oracle_transportes (clave, nombre, volumen, ESTATUS, PLACAS, RENDIMIENTO, TIPO, exportaciones) VALUES ('" + Me.txt_clave + "','" + Me.txt_nombre + "', " + Me.txt_volumen + ",'" + Me.cmb_estatus + "','" + Me.txt_placas + "'," + Me.txt_rendimiento + ",'" + Me.txt_tipo + "', " + CStr(Me.chk_exportaciones) + ")", cnn, adOpenDynamic, adLockOptimistic
                                var_modifica_registro_linea = True
                                Call pro_actualiza_ListView
                             End If
                             rs.Close
                          
                          
                          
                             strconsulta = "SELECT * FROM XXVIA_TB_TRANSPORTES WHERE CLAVE = ?"
                             With comandoORA
                                  .ActiveConnection = cnnoracle_4
                                  .CommandType = adCmdText
                                  .CommandText = strconsulta
                                  Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(Me.txt_clave))
                                  .Parameters.Append parametro
                             End With
                             Set rsaux9 = comandoORA.execute
                             Set comandoORA = Nothing
                             Set parametro = Nothing
                             If Not rsaux9.EOF Then
                                strconsulta = "UPDATE XXVIA_TB_TRANSPORTES SET NOMBRE = ?, EXPORTACIONES = ?, ESTATUS = ?, PLACAS = ?, RENDIMIENTO = ?, TIPO = ?, permsct = ?,  numpermisosct = ?, nombreaseg = ?, numpolizaseg = ?, configvehicular= ?,  placavm = ?, aniomodelovm = ?, subtiporem = ?, placa = ? WHERE CLAVE = ?"
                                With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(Me.txt_nombre))
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adNumeric, adParamInput, 20, Me.chk_exportaciones)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.cmb_estatus)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_placas)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adNumeric, adParamInput, 20, Me.txt_rendimiento)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_tipo)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_permsct)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_numpermisosct)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_nombreaseg)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_numpolizaseg)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_configvehicular)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_placaVM)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_aniomodelovm)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_subtiporem)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_placa)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(Me.txt_clave))
                                     .Parameters.Append parametro
                                End With
                                Set rsaux10 = comandoORA.execute
                                Set comandoORA = Nothing
                                Set parametro = Nothing
                     
                             Else
                  
                                 strconsulta = "INSERT INTO XXVIA_TB_TRANSPORTES (CLAVE, NOMBRE, exportaciones) VALUES (?,?, ?)"
                                 With comandoORA
                                      .ActiveConnection = cnnoracle_4
                                      .CommandType = adCmdText
                                      .CommandText = strconsulta
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(Me.txt_clave))
                                      .Parameters.Append parametro
                                      Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_nombre)
                                      .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adNumeric, adParamInput, 20, Me.chk_exportaciones)
                                     .Parameters.Append parametro
                                 End With
                                 Set rsaux10 = comandoORA.execute
                                 Set comandoORA = Nothing
                                 Set parametro = Nothing
                                 
                                strconsulta = "UPDATE XXVIA_TB_TRANSPORTES SET NOMBRE = ?, EXPORTACIONES = ?, ESTATUS = ?, PLACAS = ?, RENDIMIENTO = ?, TIPO = ?, permsct = ?,  numpermisosct = ?, nombreaseg = ?, numpolizaseg = ?, configvehicular= ?,  placavm = ?, aniomodelovm = ?, subtiporem = ?, placa = ? WHERE CLAVE = ?"
                                With comandoORA
                                     .ActiveConnection = cnnoracle_4
                                     .CommandType = adCmdText
                                     .CommandText = strconsulta
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(Me.txt_nombre))
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adNumeric, adParamInput, 20, Me.chk_exportaciones)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.cmb_estatus)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_placas)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adNumeric, adParamInput, 20, Me.txt_rendimiento)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_tipo)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_permsct)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_numpermisosct)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_nombreaseg)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_numpolizaseg)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_configvehicular)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_placaVM)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_aniomodelovm)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_subtiporem)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_placa)
                                     .Parameters.Append parametro
                                     Set parametro = .CreateParameter(, adVarChar, adParamInput, 20, CStr(Me.txt_clave))
                                     .Parameters.Append parametro
                                End With
                                Set rsaux10 = comandoORA.execute
                                Set comandoORA = Nothing
                                Set parametro = Nothing
                                 
                                 
                              End If
                          Else
                             MsgBox "Estatus incorrecto", vbOKOnly, "ATENCION"
                          End If
                       Else
                          MsgBox "Tipo de transporte incorrecto", vbOKOnly, "ATENCION"
                       End If
                    Else
                       MsgBox "Placas incorrectas", vbOKOnly, "ATENCION"
                    End If
                 Else
                    MsgBox "Rendimiento invalido", vbOKOnly, "ATENCION"
                 End If
              Else
                 MsgBox "Volumen incorrecto", vbOKOnly, "ATENCION"
              End If
           Else
              MsgBox "Nombre de transporte incorrecto", vbOKOnly, "ATENCION"
           End If
        Else
           MsgBox "Descripcion del incorrecta", vbOKOnly, "ATENCION"
        End If
End Sub


Private Sub cmd_nuevo_Click()
   Me.txt_clave = ""
   Me.txt_nombre = ""
   Me.txt_placas = ""
   Me.txt_rendimiento = ""
   Me.txt_tipo = ""
   Me.txt_volumen = ""
   Me.cmb_estatus = "ACTIVO"
   Me.txt_clave.SetFocus
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
    rs.Open "select * from tb_oracle_transportes", cnn, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
          var_clave = IIf(IsNull(rs!CLAVE), "", rs!CLAVE)
          VAR_ESTATUS = IIf(IsNull(rs!estatus), "", rs!estatus)
          var_placas = IIf(IsNull(rs!placas), "", rs!placas)
          var_rendimiento = IIf(IsNull(rs!rendimiento), 0, rs!rendimiento)
          VAR_EXPORTACIONES = IIf(IsNull(rs!EXPORTACIONES), 0, rs!EXPORTACIONES)
          var_tipo = IIf(IsNull(rs!tipo), "", rs!tipo)
          rsaux.Open "update xxvia_tb_transportes set estatus = '" + VAR_ESTATUS + "', placas = '" + var_placas + "', rendimiento = " + CStr(var_rendimiento) + ", exportaciones = " + CStr(VAR_EXPORTACIONES) + ", tipo = '" + var_tipo + "' where clave = '" + var_clave + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
          rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 4 And KeyCode = 78 Then
      cmd_nuevo_Click
   End If
   If Shift = 4 And KeyCode = 71 Then
      cmd_guardar_Click
   End If
   If Shift = 4 And KeyCode = 68 Then
   End If
   If Shift = 4 And KeyCode = 69 Then
   End If
   If Shift = 4 And KeyCode = 73 Then
   End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 0
   Left = 300
   var_modifica_registro_linea = True
   Call pro_llena_listview1
   pro_textos
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_lineas)
End Sub

Private Sub lv_transportes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_transportes, ColumnHeader)
End Sub

Private Sub lv_transportes_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Set lv_transportes.selectedItem = Item
   pro_textos
   var_modifica_registro_linea = True
End Sub

Private Sub tool_atras_siguiente_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err0:
   If Button.Index = 2 Or Button.Index = 3 Then
      lv_transportes.SetFocus
      Call pro_avanzar(Me, lv_transportes, Button)
      lv_transportes.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 1 Then
      lv_transportes.ListItems(1).Selected = True
      lv_transportes.selectedItem.EnsureVisible
      pro_textos
   End If
   If Button.Index = 4 Then
      numero_items_lineas = lv_transportes.ListItems.Count
      lv_transportes.ListItems(numero_items_lineas).Selected = True
      lv_transportes.selectedItem.EnsureVisible
      pro_textos
   End If
err0:
End Sub

Sub pro_llena_listview1()
   Dim list_item As ListItem
   rs.Open "select * from xxvia_tb_transportes", cnnoracle_4, adOpenDynamic, adLockOptimistic
   numero_items_lineas = 0
   While Not rs.EOF
      Set list_item = lv_transportes.ListItems.Add(, , rs(0).Value)
      list_item.SubItems(1) = IIf(IsNull(rs(1).Value), "", rs(1).Value)
      list_item.SubItems(2) = IIf(IsNull(rs(2).Value), "0", rs(2).Value)
      list_item.SubItems(3) = IIf(IsNull(rs!estatus), "", rs!estatus)
      list_item.SubItems(4) = IIf(IsNull(rs!placas), "", rs!placas)
      list_item.SubItems(5) = IIf(IsNull(rs!rendimiento), "0", rs!rendimiento)
      list_item.SubItems(6) = IIf(IsNull(rs!tipo), "", rs!tipo)
      list_item.SubItems(7) = IIf(IsNull(rs!EXPORTACIONES), 0, rs!EXPORTACIONES)
      
      list_item.SubItems(8) = IIf(IsNull(rs!PERMSCT), "", rs!PERMSCT)
      list_item.SubItems(9) = IIf(IsNull(rs!NUMPERMIsoSCT), "", rs!NUMPERMIsoSCT)
      list_item.SubItems(10) = IIf(IsNull(rs!NOMBREASEG), "", rs!NOMBREASEG)
      list_item.SubItems(11) = IIf(IsNull(rs!NUMPOLIZASEG), "", rs!NUMPOLIZASEG)
      list_item.SubItems(12) = IIf(IsNull(rs!configvehicular), "", rs!configvehicular)
      list_item.SubItems(13) = IIf(IsNull(rs!placavm), "", rs!placavm)
      list_item.SubItems(14) = IIf(IsNull(rs!aniomodelovm), "", rs!aniomodelovm)
      list_item.SubItems(15) = IIf(IsNull(rs!subtiporem), "", rs!subtiporem)
      list_item.SubItems(16) = IIf(IsNull(rs!PLACA), "", rs!PLACA)
      
      
      rs.MoveNext:
      numero_items_lineas = numero_items_lineas + 1
    Wend
    rs.Close
End Sub


Sub pro_textos()
On Error GoTo err0:
Dim var_n As Double
   var_n = lv_transportes.ListItems.Count
   If var_n > 0 Then
      txt_clave = lv_transportes.selectedItem
      txt_nombre = lv_transportes.selectedItem.SubItems(1)
      txt_volumen = lv_transportes.selectedItem.SubItems(2)
      Me.cmb_estatus = lv_transportes.selectedItem.SubItems(3)
      Me.txt_placas = lv_transportes.selectedItem.SubItems(4)
      Me.txt_rendimiento = lv_transportes.selectedItem.SubItems(5)
      Me.txt_tipo = lv_transportes.selectedItem.SubItems(6)
      Me.chk_exportaciones = lv_transportes.selectedItem.SubItems(7)
   
   
      Me.txt_permsct = lv_transportes.selectedItem.SubItems(8)
      Me.txt_numpermisosct = lv_transportes.selectedItem.SubItems(9)
      Me.txt_nombreaseg = lv_transportes.selectedItem.SubItems(10)
      Me.txt_numpolizaseg = lv_transportes.selectedItem.SubItems(11)
      Me.txt_configvehicular = lv_transportes.selectedItem.SubItems(12)
      Me.txt_placaVM = lv_transportes.selectedItem.SubItems(13)
      Me.txt_aniomodelovm = lv_transportes.selectedItem.SubItems(14)
      Me.txt_subtiporem = lv_transportes.selectedItem.SubItems(15)
      Me.txt_placa = lv_transportes.selectedItem.SubItems(16)
   
   End If
err0:
End Sub

Private Sub pro_actualiza_ListView()
Dim list_item As ListItem
    If var_modifica_registro_linea = True Then
       Set list_item = lv_transportes.ListItems.Add(, , txt_clave)
       list_item.SubItems(1) = txt_nombre
       If IsNumeric(Me.txt_volumen) Then
          list_item.SubItems(2) = txt_volumen
       Else
          list_item.SubItems(2) = 0
       End If
       list_item.SubItems(3) = Me.cmb_estatus
       list_item.SubItems(4) = Me.txt_placas
       list_item.SubItems(5) = Me.txt_rendimiento
       list_item.SubItems(6) = Me.txt_tipo
       list_item.SubItems(7) = Me.chk_exportaciones
       list_item.SubItems(8) = Me.txt_permsct
       list_item.SubItems(9) = Me.txt_numpermisosct
       list_item.SubItems(10) = Me.txt_nombreaseg
       list_item.SubItems(11) = Me.txt_numpolizaseg
       list_item.SubItems(12) = Me.txt_configvehicular
       list_item.SubItems(13) = Me.txt_placaVM
       list_item.SubItems(14) = Me.txt_aniomodelovm
       list_item.SubItems(15) = Me.txt_subtiporem
       list_item.SubItems(16) = Me.txt_placa
       list_item.EnsureVisible
       list_item.Selected = True
       numero_items_lineas = numero_items_lineas + 1
    Else
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).Checked = False
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index) = txt_clave
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(1) = txt_nombre
       If IsNumeric(Me.txt_volumen) Then
          lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(2) = txt_volumen
       Else
          lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(2) = 0
       End If
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(3) = Me.cmb_estatus
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(4) = Me.txt_placas
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(5) = Me.txt_rendimiento
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(6) = Me.txt_tipo
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(7) = Me.chk_exportaciones
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(8) = Me.txt_permsct
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(9) = Me.txt_numpermisosct
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(10) = Me.txt_nombreaseg
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(11) = Me.txt_numpolizaseg
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(12) = Me.txt_configvehicular
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(13) = Me.txt_placaVM
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(14) = Me.txt_aniomodelovm
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(15) = Me.txt_subtiporem
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).ListSubItems(16) = Me.txt_placa
       lv_transportes.ListItems.Item(lv_transportes.selectedItem.Index).Selected = True
    End If
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_busca_registro(Me.lv_transportes, txt_buscar, False)
      txt_buscar = ""
      pro_textos
   End If
End Sub






Private Sub txt_aseguradora_Change()

End Sub

Private Sub txt_aniomodelovm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_subtiporem.SetFocus
   End If

End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_eprmiso_SCT_Change()

End Sub

Private Sub txt_configvehicular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_placaVM.SetFocus
   End If

End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_volumen.SetFocus
   End If
End Sub

Private Sub txt_nombreaseg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_numpolizaseg.SetFocus
   End If

End Sub

Private Sub txt_numpermisosct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_nombreaseg.SetFocus
   End If
End Sub

Private Sub txt_numpolizaseg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_configvehicular.SetFocus
   End If
End Sub

Private Sub txt_permsct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_numpermisosct.SetFocus
   End If
End Sub

Private Sub txt_placa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If

End Sub

Private Sub txt_placas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_rendimiento.SetFocus
   End If
End Sub

Private Sub txt_placavm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_aniomodelovm.SetFocus
   End If

End Sub

Private Sub txt_rendimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_tipo.SetFocus
   End If
End Sub

Private Sub txt_subtiporem_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_placa.SetFocus
   End If
End Sub

Private Sub txt_tipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_permsct.SetFocus
   End If
End Sub

Private Sub txt_volumen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmb_estatus.SetFocus
   End If
End Sub
