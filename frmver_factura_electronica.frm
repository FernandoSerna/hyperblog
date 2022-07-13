VERSION 5.00
Begin VB.Form frmver_factura_electronica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir factura electronica"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   1170
      Left            =   75
      TabIndex        =   9
      Top             =   1185
      Width           =   4500
      Begin VB.TextBox txt_copias 
         Height          =   345
         Left            =   3075
         TabIndex        =   14
         Top             =   225
         Width           =   1245
      End
      Begin VB.TextBox txt_a 
         Height          =   390
         Left            =   2760
         TabIndex        =   5
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txt_de 
         Height          =   390
         Left            =   795
         TabIndex        =   4
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txt_serie 
         Height          =   390
         Left            =   795
         TabIndex        =   3
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Copias:"
         Height          =   195
         Left            =   2520
         TabIndex        =   13
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a:"
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   750
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Serie:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   405
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4215
      Picture         =   "frmver_factura_electronica.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmver_factura_electronica.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   45
      TabIndex        =   7
      Top             =   330
      Width           =   4530
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   75
      TabIndex        =   6
      Top             =   345
      Width           =   4485
      Begin VB.TextBox txt_embarque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1530
         TabIndex        =   2
         Top             =   210
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   555
         TabIndex        =   8
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmver_factura_electronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_nuevo_Click()

End Sub

Private Sub cmb_lineas_Change()

End Sub

Private Sub cmb_copias_Change()

End Sub

Private Sub cmd_imprimir_Click()
   If IsNumeric(Me.txt_copias) Then
      If IsNumeric(Me.txt_de) Then
         If IsNumeric(Me.txt_a) Then
            If CDbl(Me.txt_de) <= CDbl(Me.txt_a) Then
               If Trim(Me.txt_serie) <> "" Then
                  rs.Open "select * from tb_encabezado_cartera where vcha_ser_serie_id = '" + Me.txt_serie + "' and inte_car_numero between " + Me.txt_de + " and " + Me.txt_a + " order by inte_Car_numero", cnn, adOpenDynamic, adLockOptimistic
                  If Not rs.EOF Then
                     var_posible = 0
                     While Not rs.EOF
                           var_Archivo = var_ruta_documentos_electronicos_pdf + "\" + rs!vcha_Ser_Serie_id + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(CStr(rs!inte_Car_numero)) + ".pdf"
                           Archivoabuscar = Dir(var_Archivo)
                           If Archivoabuscar = "" Then
                              'MsgBox var_Archivo
                              var_posible = 1
                           End If
                           rs.MoveNext
                     Wend
                     rs.MoveFirst
                     If var_posible = 0 Then
                        var_si = MsgBox("Se van a imprimir las facturas de la " + Me.txt_de + " a la " + Me.txt_a, vbYesNo, "ATENCION")
                        If var_si = 6 Then
                           var_si = MsgBox("Confirmar la impresión de las facturas", vbYesNo, "ATENCION")
                           If var_si = 6 Then
                              While Not rs.EOF
                                    If (var_empresa = "02" Or var_empresa = "03") Then
                                       ccc = 0
                                       If ccc = 0 Then
                                       
                                          rsaux1.Open "select * from tb_saldos where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_Ser_serie_id = '" + Me.txt_serie + "' and inte_car_numero = " + CStr(rs!inte_Car_numero), cnn, adOpenDynamic, adLockOptimistic
                                          var_z = CDbl(Me.txt_copias)
                                          For var_j = 1 To var_z
                                              If rsaux1!FLOA_sAL_IMPORTE < 1 Then
                                                 var_Archivo = var_ruta_documentos_electronicos_pdf + "\" + Trim(rs!vcha_Ser_Serie_id) + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(Str(rs!inte_Car_numero)) + ".PDF"
                                                 'Call Shell("c:\archivos de programa\adobe\acrobat 7.0\reader\acrord32.exe  /p /h " + var_Archivo, vbMaximizedFocus)
                                                 Call Shell("c:\archivos de programa\adobe\reader 8.0\reader\acrord32.exe  /p /h " + var_Archivo, vbMaximizedFocus)
                                              Else
                                                 var_Archivo = var_ruta_documentos_electronicos_pdf + "\" + Trim(rs!vcha_Ser_Serie_id) + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(Str(rs!inte_Car_numero)) + ".PDF"
                                                 'Call Shell("c:\archivos de programa\adobe\acrobat 7.0\reader\acrord32.exe  /p /h " + var_Archivo, vbMaximizedFocus)
                                                 Call Shell("c:\archivos de programa\adobe\reader 8.0\reader\acrord32.exe  /p /h " + var_Archivo, vbMaximizedFocus)
                                              End If
                                          Next var_j
                                          rsaux1.Close
                                       Else
                                          Open (App.Path & "\EJPDF" + Trim(rs!vcha_Ser_Serie_id) + Trim(CStr(rs!inte_Car_numero)) + ".bat") For Output As #2
                                          If var_empresa = "31" Then
                                             Print #2, "START " + var_ruta_documentos_electronicos_pdf + "\" + Trim(rs!vcha_Ser_Serie_id) + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(Str(rs!inte_Car_numero)) + ".PDF"
                                          Else
                                             Print #2, "START " + var_ruta_documentos_electronicos_pdf + "\" + Trim(rs!vcha_Ser_Serie_id) + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(Str(rs!inte_Car_numero)) + ".PDF"
                                          End If
                                          Close #2
                                          var_Archivo = App.Path & "\EJPDF" + Trim(rs!vcha_Ser_Serie_id) + Trim(CStr(rs!inte_Car_numero)) + ".bat"
                                          x = Shell(var_Archivo, vbHide)
                                       End If
                                    Else
                                       rsaux1.Open "select * from tb_saldos where vcha_Emp_Empresa_id = '" + var_empresa + "' and vcha_Ser_serie_id = '" + Me.txt_serie + "' and inte_car_numero = " + CStr(rs!inte_Car_numero), cnn, adOpenDynamic, adLockOptimistic
                                       If rsaux1!FLOA_sAL_IMPORTE < 1 Then
                                          Open (App.Path & "\EJPDF" + Trim(rs!vcha_Ser_Serie_id) + Trim(CStr(rs!inte_Car_numero)) + ".bat") For Output As #2
                                          If var_empresa = "31" Then
                                             Print #2, "START " + var_ruta_documentos_electronicos_pdf + "\" + Trim(rs!vcha_Ser_Serie_id) + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(Str(rs!inte_Car_numero)) + ".PDF"
                                          Else
                                             Print #2, "START " + var_ruta_documentos_electronicos_pdf + "\" + Trim(rs!vcha_Ser_Serie_id) + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(Str(rs!inte_Car_numero)) + ".PDF"
                                          End If
                                          Close #2
                                          var_Archivo = App.Path & "\EJPDF" + Trim(rs!vcha_Ser_Serie_id) + Trim(CStr(rs!inte_Car_numero)) + ".bat"
                                          x = Shell(var_Archivo, vbHide)
                                       Else
                                          Open (App.Path & "\EJPDF" + Trim(rs!vcha_Ser_Serie_id) + Trim(CStr(rs!inte_Car_numero)) + ".bat") For Output As #2
                                          If var_empresa = "31" Then
                                             Print #2, "START " + var_ruta_documentos_electronicos_pdf + "\" + Trim(rs!vcha_Ser_Serie_id) + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(Str(rs!inte_Car_numero)) + ".PDF"
                                          Else
                                             Print #2, "START " + var_ruta_documentos_electronicos_pdf + "\" + Trim(rs!vcha_Ser_Serie_id) + "\" + Trim(rs!vcha_Ser_Serie_id) + Trim(Str(rs!inte_Car_numero)) + ".PDF"
                                          End If
                                          Close #2
                                          var_Archivo = App.Path & "\EJPDF" + Trim(rs!vcha_Ser_Serie_id) + Trim(CStr(rs!inte_Car_numero)) + ".bat"
                                          x = Shell(var_Archivo, vbHide)
                                       End If
                                       rsaux1.Close
                                    End If
                                    rs.MoveNext
                              Wend
                           Else
                              MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
                           End If
                        Else
                           MsgBox "Se a cancelado la impresión de las facturas", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "Aun no se generan todas las facturas electronicas, favor de esperar unos minutos mas", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Las facturas no existen", vbOKOnly, "ATENCION"
                  End If
                  rs.Close
               Else
                  MsgBox "Debe de seleccionar una serie", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "La factura inicial no puede ser mayor a la factura final", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Número de factura final incorrecto", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Número de factura inicial incorrecto", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Número de copias incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   'MsgBox var_empresa
   If var_empresa = "15" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sid;Data Source=admcdindustrial"
   End If
   If var_empresa = "16" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sid;Data Source=admcdindustrial"
   End If
   If var_empresa = "30" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
   End If
   If var_empresa = "31" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sidcantia;Data Source=sqlquezada2"
   End If
   If var_empresa = "38" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
   End If
   If var_empresa = "02" Or var_empresa = "03" Or var_empresa = "37" Or var_empresa = "29" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=vianney;Data Source=DISTRIBUCION"
   End If
   If var_empresa = "06" Or var_empresa = "17" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sid;Data Source=admcdindustrial"
   End If
   If var_empresa = "18" Then
      var_conexion = "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=sidtextilera;Data Source=sqlquezada2"
   End If
   If cnn_ver_factura_electronica.State = 1 Then
      cnn_ver_factura_electronica.Close
   End If
   'MsgBox cnn_ver_factura_electronica.ConnectionString
   'MsgBox var_empresa
   cnn_ver_factura_electronica.Open var_conexion
   If var_empresa = "02" Or var_empresa = "03" Then
      Me.txt_copias = 2
   Else
      Me.txt_copias = 1
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub txt_a_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_de_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_a.SetFocus
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_embarque) Then
         var_cadena = "SELECT  dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID, MIN(dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO) AS MINIMO, MAX(dbo.TB_ENCABEZADO_CARTERA.INTE_CAR_NUMERO) As MAXIMO FROM dbo.TB_ENCABEZADO_EMBARQUES INNER JOIN dbo.TB_DETALLE_EMBARQUES ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID AND dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = dbo.TB_DETALLE_EMBARQUES.INTE_EMB_EMBARQUE INNER JOIN dbo.TB_ENCABEZADO_CARTERA ON dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_EMP_EMPRESA_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_EMP_EMPRESA_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_UOR_UNIDAD_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_UOR_UNIDAD_ID AND dbo.TB_DETALLE_EMBARQUES.VCHA_MOV_MOVIMIENTO_ID = dbo.TB_ENCABEZADO_CARTERA.VCHA_MOV_MOVIMIENTO_ID AND dbo.TB_DETALLE_EMBARQUES.INTE_SAL_NUMERO = dbo.TB_ENCABEZADO_CARTERA.INTE_EMO_NUMERO "
         var_cadena = var_cadena + " WHERE (dbo.TB_ENCABEZADO_EMBARQUES.INTE_EMB_EMBARQUE = " + Me.txt_embarque + ") AND (dbo.TB_ENCABEZADO_EMBARQUES.VCHA_EMP_EMPRESA_ID = '" + var_empresa + "') AND (dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS <> 'C' OR dbo.TB_ENCABEZADO_CARTERA.CHAR_CAR_ESTATUS IS NULL)  GROUP BY dbo.TB_ENCABEZADO_CARTERA.VCHA_SER_SERIE_ID"
         'MsgBox cnn_ver_factura_electronica
         rs.Open var_cadena, cnn_ver_factura_electronica, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_serie = rs!vcha_Ser_Serie_id
            Me.txt_de = rs!MINIMO
            Me.txt_a = rs!maximo
         Else
            Me.txt_serie = ""
            Me.txt_de = ""
            Me.txt_a = ""
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         Me.txt_serie = ""
         Me.txt_de = ""
         Me.txt_a = ""
         MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
      End If
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_serie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_de.SetFocus
   End If
End Sub
