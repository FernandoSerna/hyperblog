VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdescuentos_promociones_textilera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descuentos de Promociones"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      Picture         =   "frmdescuentos_promociones_textilera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Caption         =   "11"
      Height          =   315
      Left            =   765
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Aplicar Promociones con codigos de 11 digitos"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_cargar_promociones 
      Appearance      =   0  'Flat
      Caption         =   "10"
      Height          =   315
      Left            =   435
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Aplicar Promociones a codigos con 10 digitos"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_genera_tabla 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmdescuentos_promociones_textilera.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Generar Tabla de Promociones para Tiendas"
      Top             =   15
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   90
      TabIndex        =   18
      Top             =   4980
      Width           =   5685
      Begin VB.TextBox txt_oferta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4290
         TabIndex        =   5
         Top             =   240
         Width           =   1260
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1245
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Oferta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   20
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   19
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   840
      Left            =   105
      TabIndex        =   15
      Top             =   4080
      Width           =   5670
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1035
         TabIndex        =   17
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   3285
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5370
      Picture         =   "frmdescuentos_promociones_textilera.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmdescuentos_promociones_textilera.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Aplicar Promociones"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame5 
      Height          =   30
      Left            =   30
      TabIndex        =   12
      Top             =   345
      Width           =   5715
   End
   Begin VB.Frame Frame2 
      Caption         =   " Canales de Venta "
      Height          =   3630
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   5685
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmdescuentos_promociones_textilera.frx":0988
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmdescuentos_promociones_textilera.frx":0B9E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Marcar (Enter)"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmdescuentos_promociones_textilera.frx":0DE8
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   105
         Picture         =   "frmdescuentos_promociones_textilera.frx":0EBA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmdescuentos_promociones_textilera.frx":0FBC
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   240
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_agentes 
         Height          =   2835
         Left            =   45
         TabIndex        =   1
         Top             =   720
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   5001
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
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "marca"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   30
         TabIndex        =   10
         Top             =   525
         Width           =   5610
      End
   End
End
Attribute VB_Name = "frmdescuentos_promociones_textilera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_cargar_promociones_Click()
   Dim VERIFICADOR As Integer
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   Dim var_descuento_string As String
   Dim strConnectionString As String
   Dim recordSet As New ADODB.recordSet
   'On Error GoTo salir:
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & "C:\Ofertas.XLS"
   recordSet.Open "SELECT * FROM [Hoja1$]", strConnectionString
   rs.Open "delete from ofertas", cnn, adOpenDynamic, adLockOptimistic
   While Not recordSet.EOF
         If IsNull(recordSet!codigo) Then
         
         Else
            rs.Open "insert into ofertas (codigo, descuento, vcha_art_articulo_id) values ('" + CStr(recordSet!codigo) + "'," + CStr(recordSet!descuento) + ",'" + CStr(recordSet!codigo) + "')", cnn, adOpenDynamic, adLockOptimistic
         End If
         recordSet.MoveNext
   Wend
   If IsDate(Me.txt_inicio) Then
            If IsDate(Me.txt_fin) Then
               If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
                  VAR_CADENA_CANALES = ""
            
                  var_dia = CStr(Day(CDate(txt_inicio)))
                  var_mes = CStr(Month(CDate(txt_inicio)))
                  var_año = CStr(Year(CDate(txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                 
               
                  var_dia = CStr(Day(CDate(txt_fin)))
                  var_mes = CStr(Month(CDate(txt_fin)))
                  var_año = CStr(Year(CDate(txt_fin)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
                  
               
                  For var_i = 1 To lv_agentes.ListItems.Count
                      lv_agentes.ListItems.Item(var_i).Selected = True
                      If lv_agentes.selectedItem.SubItems(2) = "*" Then
                             rs.Open "SELECT * FROM ofertas", cnn, adOpenDynamic, adLockOptimistic
                             While Not rs.EOF
                                   rsaux5.Open "update ofertas set marca = " + CStr(var_i) + " where vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                   var_codigo_inicio = Mid(rs!vcha_Art_articulo_id, 1, 10)
                                   var_z = CInt(Mid(rs!vcha_Art_articulo_id, 10, 1))
                                   Me.txt_oferta = IIf(IsNull(rs!descuento), 0, rs!descuento)
                                   For var_z = 0 To 9
                                        var_codigo = Trim(var_codigo_inicio) + Trim(CStr(var_z))
                                        sum1 = 0
                                        sum2 = 0
                                        mcodigo = var_codigo
                                        longitud = Len(mcodigo)
                                        For icont = 1 To longitud
                                            If ((icont / 2) - Int((icont / 2))) = 0 Then
                                               sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                                            Else
                                               sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                                            End If
                                        Next icont
                                        msuma = sum1 * 13 + sum2
                                        VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                                        If VERIFICADOR = 10 Then
                                           VERIFICADOR = 0
                                        End If
                                        var_codigo = var_codigo + Trim(CStr(VERIFICADOR))
                                        
                                        rsaux.Open "SELECT * FROM TB_DESCUENTOS_PROMOCIONES WHERE VCHA_CAN_CANAL_VENTA_ID = '" + Trim(lv_agentes.selectedItem) + "' AND VCHA_ART_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                        If rsaux.EOF Then
                                           rsaux2.Open "insert into tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + lv_agentes.selectedItem + "', '" + rs!vcha_Art_articulo_id + "'," + var_fecha_inicio + ", " + var_fecha_fin + "," + Me.txt_oferta + ")", cnn, adOpenDynamic, adLockOptimistic
                                        Else
                                           rsaux2.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = " + var_fecha_inicio + ", dtim_dpr_fecha_fin = " + var_fecha_fin + ", floa_dpr_descuento = " + Me.txt_oferta + " where vcha_can_canal_venta_id = '" + Me.lv_agentes.selectedItem + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        rsaux.Close
                                        var_descuento_string = Mid(var_codigo, 11, 1)
                                        var_cadena = "SELECT * FROM TB_PROMOCIONES_TIENDAS_TEXTILERA WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'"
                                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                        If rsaux.EOF Then
                                           rsaux2.Open "Insert Into TB_PROMOCIONES_TIENDAS_TEXTILERA (vcha_can_Canal_venta_id, codigo, descuento0, descuento1, descuento2, descuento3, descuento4, descuento5, descuento6, descuento7, descuento8, descuento9, vigencia_i, vigencia_f, cantidad) values ('" + Me.lv_agentes.selectedItem + "', '" + Trim(Mid(var_codigo, 1, 10)) + "',0,0,0,0,0,0,0,0,0,0," + var_fecha_inicio + "," + var_fecha_fin + ",0)", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "0" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento0 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "1" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento1 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "2" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento2 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "3" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento3 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "4" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento4 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "5" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento5 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "6" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento6 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "7" Then
                                            rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento7 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "8" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento8 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "9" Then
                                            rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento9 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        rsaux.Close
                                   Next var_z
                                   rs.MoveNext
                             Wend
                             rs.Close
                       End If
                  Next var_i
                  MsgBox "Se a terminado de actualizar las promociones", vbOKOnly, "ATENCION"
               Else
                  MsgBox "La fecha de inicio debe de ser menor a la final", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
         End If
Exit Sub
salir:
   MsgBox "A surgido un error al cargar las ofertas, verifique que el archivo de ofertas.xls este correcto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
End Sub

Private Sub cmd_genera_tabla_Click()
   Dim var_inicio As String
   Dim var_final As String
   var_si = MsgBox("¿Se generara la tabla de promociones para tiendas?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      var_inicio = CStr(Now)
      rs.Open "EXEC SP_PROMOCIONES_TIENDAS_TEXTILERAS", cnn, adOpenDynamic, adLockOptimistic
      var_fin = CStr(Now)
      MsgBox "Se termino la creación de la tabla de promociones para tiendas, inicio " + var_inicio + ", termino " + CStr(var_fin), vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_imprimir_Click()
   Dim VERIFICADOR As Integer
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   Dim var_descuento_string As String
   If IsNumeric(Me.txt_oferta) Then
      If CDbl(Me.txt_oferta) <= 100 Then
         If IsDate(Me.txt_inicio) Then
            If IsDate(Me.txt_fin) Then
               If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
                  VAR_CADENA_CANALES = ""
            
                  var_dia = CStr(Day(CDate(txt_inicio)))
                  var_mes = CStr(Month(CDate(txt_inicio)))
                  var_año = CStr(Year(CDate(txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                 
               
                  var_dia = CStr(Day(CDate(txt_fin)))
                  var_mes = CStr(Month(CDate(txt_fin)))
                  var_año = CStr(Year(CDate(txt_fin)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
               
               
                  For var_i = 1 To lv_agentes.ListItems.Count
                      lv_agentes.ListItems.Item(var_i).Selected = True
                      If lv_agentes.selectedItem.SubItems(2) = "*" Then
                          If Len(Trim(Me.txt_codigo)) <= 10 Then
                             rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID LIKE '" + Trim(Me.txt_codigo) + "%' and substring(vcha_art_articulo_id,11,1) = '0'", cnn, adOpenDynamic, adLockOptimistic
                             While Not rs.EOF
                                   var_codigo_1 = Mid(rs!vcha_Art_articulo_id, 1, 10)
                                   For var_j = 0 To 9
                                       var_codigo = var_codigo_1 + Trim(CStr(var_j))
                                       sum1 = 0
                                       sum2 = 0
                                       mcodigo = var_codigo
                                       longitud = Len(mcodigo)
                                       For icont = 1 To longitud
                                           If ((icont / 2) - Int((icont / 2))) = 0 Then
                                              sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                                           Else
                                              sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                                           End If
                                       Next icont
                                       msuma = sum1 * 13 + sum2
                                       VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                                       If VERIFICADOR = 10 Then
                                          VERIFICADOR = 0
                                       End If
                                       var_codigo = var_codigo + Trim(CStr(VERIFICADOR))
                 
                                       rsaux.Open "SELECT * FROM TB_DESCUENTOS_PROMOCIONES WHERE VCHA_CAN_CANAL_VENTA_ID = '" + Trim(lv_agentes.selectedItem) + "' AND VCHA_ART_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If rsaux.EOF Then
                                          rsaux2.Open "insert into tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + lv_agentes.selectedItem + "', '" + var_codigo + "'," + var_fecha_inicio + ", " + var_fecha_fin + "," + Me.txt_oferta + ")", cnn, adOpenDynamic, adLockOptimistic
                                       Else
                                          rsaux2.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = " + var_fecha_inicio + ", dtim_dpr_fecha_fin = " + var_fecha_fin + ", floa_dpr_descuento = " + Me.txt_oferta + " where vcha_can_canal_venta_id = '" + Me.lv_agentes.selectedItem + "' and vcha_art_articulo_id = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       rsaux.Close
                                       var_descuento_string = Mid(var_codigo, 11, 1)
                                       var_codigo_text_ = Mid(var_codigo, 1, 10)
                                       var_cadena = "SELECT * FROM TB_PROMOCIONES_TIENDAS_TEXTILERA WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'"
                                       rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                       If rsaux.EOF Then
                                          rsaux2.Open "Insert Into TB_PROMOCIONES_TIENDAS_TEXTILERA (vcha_can_Canal_venta_id, codigo, descuento0, descuento1, descuento2, descuento3, descuento4, descuento5, descuento6, descuento7, descuento8, descuento9, vigencia_i, vigencia_f, cantidad) values ('" + Me.lv_agentes.selectedItem + "', '" + Trim(Mid(var_codigo, 1, 10)) + "',0,0,0,0,0,0,0,0,0,0," + var_fecha_inicio + "," + var_fecha_fin + ",0)", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "0" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento0 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "1" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento1 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "2" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento2 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "3" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento3 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "4" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento4 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "5" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento5 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "6" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento6 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "7" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento7 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "8" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento8 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If var_descuento_string = "9" Then
                                          rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento9 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       rsaux.Close
                                   Next var_j
                                   rs.MoveNext
                             Wend
                             rs.Close
                          Else
                             rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID LIKE '" + Trim(Me.txt_codigo) + "%'", cnn, adOpenDynamic, adLockOptimistic
                             While Not rs.EOF
                                   rsaux.Open "SELECT * FROM TB_DESCUENTOS_PROMOCIONES WHERE VCHA_CAN_CANAL_VENTA_ID = '" + Trim(lv_agentes.selectedItem) + "' AND VCHA_ART_ARTICULO_ID = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                   If rsaux.EOF Then
                                      rsaux2.Open "insert into tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + lv_agentes.selectedItem + "', '" + rs!vcha_Art_articulo_id + "'," + var_fecha_inicio + ", " + var_fecha_fin + "," + Me.txt_oferta + ")", cnn, adOpenDynamic, adLockOptimistic
                                   Else
                                      rsaux2.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = " + var_fecha_inicio + ", dtim_dpr_fecha_fin = " + var_fecha_fin + ", floa_dpr_descuento = " + Me.txt_oferta + " where vcha_can_canal_venta_id = '" + Me.lv_agentes.selectedItem + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   rsaux.Close
                                   var_codigo = rs!vcha_Art_articulo_id
                                   var_descuento_string = Mid(var_codigo, 11, 1)
                                   var_cadena = "SELECT * FROM TB_PROMOCIONES_TIENDAS_TEXTILERA WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'"
                                   rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                   If rsaux.EOF Then
                                      rsaux2.Open "Insert Into TB_PROMOCIONES_TIENDAS_TEXTILERA (vcha_can_Canal_venta_id, codigo, descuento0, descuento1, descuento2, descuento3, descuento4, descuento5, descuento6, descuento7, descuento8, descuento9, vigencia_i, vigencia_f, cantidad) values ('" + Me.lv_agentes.selectedItem + "', '" + Trim(Mid(var_codigo, 1, 10)) + "',0,0,0,0,0,0,0,0,0,0," + var_fecha_inicio + "," + var_fecha_fin + ",0)", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "0" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento0 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "1" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento1 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "2" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento2 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "3" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento3 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "4" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento4 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "5" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento5 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "6" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento6 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "7" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento7 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "8" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento8 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   If var_descuento_string = "9" Then
                                      rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento9 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   rsaux.Close
                                   
                                   rs.MoveNext
                             Wend
                             rs.Close
                          End If
                       End If
                  Next var_i
                  MsgBox "Se a terminado de actualizar las promociones", vbOKOnly, "ATENCION"
               Else
                  MsgBox "La fecha de inicio debe de ser menor a la final", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "La oferta no debe de ser mayor al 100%", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Oferta incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   n = lv_agentes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_agentes.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_agentes.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command2_Click()
   i = lv_agentes.selectedItem.Index
   If lv_agentes.selectedItem.SubItems(2) = "*" Then
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_agentes.Refresh
   Else
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_agentes.Refresh
   End If
End Sub

Private Sub Command3_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
         lv_agentes.selectedItem.SubItems(2) = ""
         lv_agentes.ListItems.Item(i).Bold = False
         lv_agentes.ListItems.Item(i).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command4_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = ""
      lv_agentes.ListItems.Item(i).Bold = False
      lv_agentes.ListItems.Item(i).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_agentes.Refresh
End Sub

Private Sub Command5_Click()
   n = lv_agentes.ListItems.Count
   For i = 1 To n
      lv_agentes.ListItems.Item(i).Selected = True
      lv_agentes.selectedItem.SubItems(2) = "*"
      lv_agentes.ListItems.Item(i).Bold = True
      lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_agentes.Refresh
End Sub

Private Sub Command6_Click()
   Dim VERIFICADOR As Integer
   Dim var_fecha_fin_1 As Date
   Dim dia As String
   Dim mes As String
   Dim año As String
   Dim var_fecha_inicio As String
   Dim var_fecha_fin As String
   Dim var_consecutivo As Integer
   Dim var_afectacion_movimiento As String
   Dim var_vistas As String
   Dim var_descuento_string As String
   Dim strConnectionString As String
   Dim recordSet As New ADODB.recordSet
   'On Error GoTo salir:
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & "C:\Ofertas.XLS"
   recordSet.Open "SELECT * FROM [Hoja1$]", strConnectionString
   rs.Open "delete from ofertas", cnn, adOpenDynamic, adLockOptimistic
   While Not recordSet.EOF
         rs.Open "insert into ofertas (codigo, descuento, vcha_art_articulo_id) values ('" + CStr(IIf(IsNull(recordSet!codigo), "", recordSet!codigo)) + "'," + CStr(IIf(IsNull(recordSet!descuento), 0, recordSet!descuento)) + ",'" + CStr(IIf(IsNull(recordSet!codigo), "", recordSet!codigo)) + "')", cnn, adOpenDynamic, adLockOptimistic
         recordSet.MoveNext
   Wend
   If IsDate(Me.txt_inicio) Then
            If IsDate(Me.txt_fin) Then
               If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
                  VAR_CADENA_CANALES = ""
            
                  var_dia = CStr(Day(CDate(txt_inicio)))
                  var_mes = CStr(Month(CDate(txt_inicio)))
                  var_año = CStr(Year(CDate(txt_inicio)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                 
               
                  var_dia = CStr(Day(CDate(txt_fin)))
                  var_mes = CStr(Month(CDate(txt_fin)))
                  var_año = CStr(Year(CDate(txt_fin)))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
               
                  
               
                  For var_i = 1 To lv_agentes.ListItems.Count
                      lv_agentes.ListItems.Item(var_i).Selected = True
                      If lv_agentes.selectedItem.SubItems(2) = "*" Then
                             rs.Open "SELECT * FROM ofertas", cnn, adOpenDynamic, adLockOptimistic
                             While Not rs.EOF
                                   rsaux5.Open "update ofertas set marca = " + CStr(var_i) + " where vcha_Art_Articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                   var_codigo_inicio = Mid(rs!vcha_Art_articulo_id, 1, 10)
                                   If rs!vcha_Art_articulo_id <> "" Then
                                   var_z = CInt(Mid(rs!vcha_Art_articulo_id, 11, 1))
                                   Me.txt_oferta = IIf(IsNull(rs!descuento), 0, rs!descuento)
                                   'For var_z = 0 To 9
                                        var_codigo = Trim(var_codigo_inicio) + Trim(CStr(var_z))
                                        sum1 = 0
                                        sum2 = 0
                                        mcodigo = var_codigo
                                        longitud = Len(mcodigo)
                                        For icont = 1 To longitud
                                            If ((icont / 2) - Int((icont / 2))) = 0 Then
                                               sum2 = sum2 + Val(Mid(mcodigo, icont, 1))
                                            Else
                                               sum1 = sum1 + Val(Mid(mcodigo, icont, 1))
                                            End If
                                        Next icont
                                        msuma = sum1 * 13 + sum2
                                        VERIFICADOR = 10 - ((msuma / 10) - Int(msuma / 10)) * 10
                                        If VERIFICADOR = 10 Then
                                           VERIFICADOR = 0
                                        End If
                                        var_codigo = var_codigo + Trim(CStr(VERIFICADOR))
                                        
                                        rsaux.Open "SELECT * FROM TB_DESCUENTOS_PROMOCIONES WHERE VCHA_CAN_CANAL_VENTA_ID = '" + Trim(lv_agentes.selectedItem) + "' AND VCHA_ART_ARTICULO_ID = '" + var_codigo + "'", cnn, adOpenDynamic, adLockOptimistic
                                        If rsaux.EOF Then
                                           rsaux2.Open "insert into tb_descuentos_promociones (vcha_can_canal_venta_id, vcha_art_articulo_id, dtim_dpr_fecha_inicio, dtim_dpr_fecha_fin, floa_dpr_descuento) values ('" + lv_agentes.selectedItem + "', '" + rs!vcha_Art_articulo_id + "'," + var_fecha_inicio + ", " + var_fecha_fin + "," + Me.txt_oferta + ")", cnn, adOpenDynamic, adLockOptimistic
                                        Else
                                           rsaux2.Open "update tb_descuentos_promociones set dtim_dpr_fecha_inicio = " + var_fecha_inicio + ", dtim_dpr_fecha_fin = " + var_fecha_fin + ", floa_dpr_descuento = " + Me.txt_oferta + " where vcha_can_canal_venta_id = '" + Me.lv_agentes.selectedItem + "' and vcha_art_articulo_id = '" + rs!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        rsaux.Close
                                        
                                        var_descuento_string = Mid(var_codigo, 11, 1)
                                        var_cadena = "SELECT * FROM TB_PROMOCIONES_TIENDAS_TEXTILERA WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'"
                                        rsaux.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                                        If rsaux.EOF Then
                                           rsaux2.Open "Insert Into TB_PROMOCIONES_TIENDAS_TEXTILERA (vcha_can_Canal_venta_id, codigo, descuento0, descuento1, descuento2, descuento3, descuento4, descuento5, descuento6, descuento7, descuento8, descuento9, vigencia_i, vigencia_f, cantidad) values ('" + Me.lv_agentes.selectedItem + "', '" + Trim(Mid(var_codigo, 1, 10)) + "',0,0,0,0,0,0,0,0,0,0," + var_fecha_inicio + "," + var_fecha_fin + ",0)", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "0" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento0 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "1" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento1 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "2" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento2 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "3" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento3 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "4" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento4 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "5" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento5 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "6" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento6 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "7" Then
                                            rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento7 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "8" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento8 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        If var_descuento_string = "9" Then
                                           rsaux2.Open "update TB_PROMOCIONES_TIENDAS_TEXTILERA set descuento9 = " + Me.txt_oferta + ", vigencia_i = " + var_fecha_inicio + ", vigencia_f = " + var_fecha_fin + " WHERE CODIGO = '" + Mid(var_codigo, 1, 10) + "' and vcha_can_canal_venta_id = '" + Trim(lv_agentes.selectedItem) + "'", cnn, adOpenDynamic, adLockOptimistic
                                        End If
                                        rsaux.Close
                                      'Next var_z
                                   End If
                                   rs.MoveNext
                             Wend
                             rs.Close
                       End If
                  Next var_i
                  MsgBox "Se a terminado de actualizar las promociones", vbOKOnly, "ATENCION"
               Else
                  MsgBox "La fecha de inicio debe de ser menor a la final", vbOKOnly, "ATENCION"
               End If
            Else
               MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
            End If
         Else
            MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
         End If
Exit Sub
salir:
   MsgBox "A surgido un error al cargar las ofertas, verifique que el archivo de ofertas.xls este correcto", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
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
End Sub

Private Sub Command7_Click()
      
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
            
            
      var_dia = CStr(Day(Date + 1))
      var_mes = CStr(Month(Date + 1))
      var_año = CStr(Year(Date + 1))
      If Len(Trim(var_dia)) = 1 Then
         var_dia = "0" + var_dia
      End If
      If Len(Trim(var_mes)) = 1 Then
         var_mes = "0" + var_mes
      End If
      var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"


      Set reporte = appl.OpenReport(App.Path + "\rep_descuentos_promociones.rpt")
      reporte.RecordSelectionFormula = "{VW_REPORTE_DESCUENTOS_PROMOCIONES.DTIM_DPR_FECHA_FIN} >= cdate('" + CStr(Date) + "')"
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de Antigüedad de Saldos Historico"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("¿Desea exportar el reporte a excel?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_descuentos_promociones.rpt")
         reporte.RecordSelectionFormula = "{VW_REPORTE_DESCUENTOS_PROMOCIONES.DTIM_DPR_FECHA_FIN} >= cdate('" + CStr(Date) + "')"
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reportessid\" + var_nombre_empresa + "_REPORTE_ANTIGuedad_saldos_" + Me.lv_agentes.selectedItem + "_" + Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
      End If
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 1000
   Left = 3000
   Me.txt_inicio = Date
   Me.txt_fin = Date
   rs.Open "SELECT VCHA_CAN_CANAL_VENTA_ID, VCHA_CAN_NOMBRE From VW_CLIENTES WHERE (VCHA_TIT_TITULAR_ID = 'T000001423') order by vcha_can_nombre", cnn, adOpenDynamic, adLockOptimistic
   numero_items_agentes = 0
   While Not rs.EOF
      Set list_item = lv_agentes.ListItems.Add(, , rs!vcha_can_canal_venta_id)
      list_item.SubItems(1) = IIf(IsNull(rs!vcha_can_nombre), "", rs!vcha_can_nombre)
      list_item.SubItems(2) = ""
      rs.MoveNext:
      numero_items_agentes = numero_items_agentes + 1
   Wend
   rs.Close
   If numero_items_agentes > 12 Then
      lv_agentes.ColumnHeaders(2).Width = 4200.71
   Else
      lv_agentes.ColumnHeaders(2).Width = 4499.71
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_articulos2)
End Sub

Private Sub lv_agentes_BeforeLabelEdit(Cancel As Integer)
   Call pro_ordena_listas(lv_agentes, ColumnHeader)
End Sub

Private Sub lv_agentes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_agentes.selectedItem.Index
      If lv_agentes.selectedItem.SubItems(2) = "*" Then
         lv_agentes.selectedItem.SubItems(2) = ""
         lv_agentes.ListItems.Item(i).Bold = False
         lv_agentes.ListItems.Item(i).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_agentes.Refresh
      Else
         lv_agentes.selectedItem.SubItems(2) = "*"
         lv_agentes.ListItems.Item(i).Bold = True
         lv_agentes.ListItems.Item(i).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_agentes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_agentes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_agentes.Refresh
      End If
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub

Private Sub txt_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar la fecha"
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_oferta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub
