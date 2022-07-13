VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdesuento_promocion_cliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descuentos de promoción por cliente"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_carga_masiva 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      Picture         =   "frmdesuento_promocion_cliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Carga Masiva"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Caption         =   " Vigencia "
      Height          =   720
      Left            =   120
      TabIndex        =   22
      Top             =   4635
      Width           =   7500
      Begin VB.TextBox txt_fin 
         Height          =   360
         Left            =   4575
         TabIndex        =   2
         Top             =   255
         Width           =   1275
      End
      Begin VB.TextBox txt_inicio 
         Height          =   360
         Left            =   2025
         TabIndex        =   1
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   4260
         TabIndex        =   24
         Top             =   345
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   1560
         TabIndex        =   23
         Top             =   345
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Clientes "
      Height          =   4125
      Left            =   120
      TabIndex        =   14
      Top             =   465
      Width           =   7500
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   435
         Picture         =   "frmdesuento_promocion_cliente.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Marcar Todos Alt + T"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   105
         Picture         =   "frmdesuento_promocion_cliente.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Desmarcar Todos Alt + D"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1095
         Picture         =   "frmdesuento_promocion_cliente.frx":041A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Invertir Selección Alt + V"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         Picture         =   "frmdesuento_promocion_cliente.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Marcar (Enter)"
         Top             =   240
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1425
         Picture         =   "frmdesuento_promocion_cliente.frx":0736
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Marcar Rango Alt + R"
         Top             =   240
         Width           =   330
      End
      Begin MSComctlLib.ListView lv_clientes 
         Height          =   3345
         Left            =   45
         TabIndex        =   20
         Top             =   720
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   5900
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   9878
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
         TabIndex        =   21
         Top             =   525
         Width           =   7440
      End
   End
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      Picture         =   "frmdesuento_promocion_cliente.frx":094C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Eliminar Alt + E"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7320
      Picture         =   "frmdesuento_promocion_cliente.frx":0A4E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   420
      Picture         =   "frmdesuento_promocion_cliente.frx":1088
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmdesuento_promocion_cliente.frx":118A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   75
      TabIndex        =   6
      Top             =   315
      Width           =   7635
   End
   Begin VB.Frame Frame1 
      Caption         =   " Promociones "
      Height          =   1560
      Left            =   135
      TabIndex        =   0
      Top             =   5385
      Width           =   7485
      Begin VB.CheckBox chk_marca 
         Caption         =   "Tomar en cuenta el descuento del cliente"
         Height          =   450
         Left            =   1005
         TabIndex        =   13
         Top             =   1065
         Width           =   3600
      End
      Begin VB.TextBox txt_descuento 
         Height          =   350
         Left            =   990
         TabIndex        =   5
         Top             =   690
         Width           =   1470
      End
      Begin VB.TextBox txt_nombre_articulo 
         Height          =   350
         Left            =   2490
         TabIndex        =   4
         Top             =   315
         Width           =   4905
      End
      Begin VB.TextBox txt_articulo 
         Height          =   350
         Left            =   990
         TabIndex        =   3
         Top             =   315
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descuento:"
         Height          =   195
         Left            =   75
         TabIndex        =   11
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Articulo:"
         Height          =   195
         Left            =   75
         TabIndex        =   10
         Top             =   390
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmdesuento_promocion_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_carga_masiva_Click()
   Dim var_marca As String
   Dim VERIFICADOR As Integer
   Dim strConnectionString As String
   Dim recordSet As New ADODB.recordSet
   On Error GoTo salir:
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            var_dia = Day(Me.txt_inicio)
            var_mes = Month(Me.txt_inicio)
            var_año = Year(Me.txt_inicio)
            If Len(CStr(var_dia)) = 1 Then
               var_dia_str = "0" + CStr(var_dia)
            Else
               var_dia_str = CStr(var_dia)
            End If
            If Len(CStr(var_mes)) = 1 Then
               var_mes_str = "0" + CStr(var_mes)
            Else
               var_mes_str = CStr(var_mes)
            End If
            If Len(CStr(var_año)) = 1 Then
               var_año_str = "200" + CStr(var_año)
            Else
               If Len(CStr(var_año)) = 2 Then
                  var_año_str = "20" + CStr(var_año)
               Else
                  If Len(CStr(var_año)) = 3 Then
                     var_año_str = "2" + CStr(var_año)
                  Else
                     var_año_str = CStr(var_año)
                  End If
               End If
            End If
            var_fecha_inicio = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
            
            var_dia = Day(Me.txt_fin)
            var_mes = Month(Me.txt_fin)
            var_año = Year(Me.txt_fin)
            If Len(CStr(var_dia)) = 1 Then
               var_dia_str = "0" + CStr(var_dia)
            Else
               var_dia_str = CStr(var_dia)
            End If
            If Len(CStr(var_mes)) = 1 Then
               var_mes_str = "0" + CStr(var_mes)
            Else
               var_mes_str = CStr(var_mes)
            End If
            If Len(CStr(var_año)) = 1 Then
               var_año_str = "200" + CStr(var_año)
            Else
               If Len(CStr(var_año)) = 2 Then
                  var_año_str = "20" + CStr(var_año)
               Else
                  If Len(CStr(var_año)) = 3 Then
                     var_año_str = "2" + CStr(var_año)
                  Else
                     var_año_str = CStr(var_año)
                  End If
               End If
            End If
            var_fecha_fin = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
            
            strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & "C:\Ofertas.XLS"
            recordSet.Open "SELECT * FROM [Hoja1$]", strConnectionString
            rs.Open "delete from ofertas", cnn, adOpenDynamic, adLockOptimistic
            While Not recordSet.EOF
                  If IsNull(recordSet!codigo) Then
                  Else
                     rs.Open "insert into ofertas (codigo, descuento, vcha_art_articulo_id) values ('" + CStr(recordSet!codigo) + "'," + CStr(IIf(IsNull(recordSet!descuento), 0, recordSet!descuento)) + ",'" + CStr(recordSet!codigo) + "')", cnn, adOpenDynamic, adLockOptimistic
                  End If
                  recordSet.MoveNext
            Wend
            recordSet.Close
            rsaux10.Open "select * from ofertas", cnn, adOpenDynamic, adLockOptimistic
            While Not rsaux10.EOF
                  Me.txt_Articulo = IIf(IsNull(rsaux10!vcha_Art_articulo_id), "", rsaux10!vcha_Art_articulo_id)
                  Me.txt_nombre_articulo = ""
                  If Trim(Me.txt_Articulo) <> "" Then
                     If Len(Me.txt_Articulo) = 10 Then
                        rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID LIKE  '" + Me.txt_Articulo + "%'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rs.EOF Then
                           Me.txt_nombre_articulo = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
                        Else
                           Me.txt_Articulo = ""
                           Me.txt_nombre_articulo = ""
                        End If
                        rs.Close
                     Else
                        If Len(Trim(Me.txt_Articulo)) = 5 Then
                           If rsaux5.State = 1 Then
                              rsaux5.Close
                           End If
                           rsaux5.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + Mid(Me.txt_Articulo, 1, 1) + "' and vcha_div_division_id = '" + Mid(Me.txt_Articulo, 2, 2) + "' and vcha_sub_subdivision_id = '" + Mid(Me.txt_Articulo, 4, 2) + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux5.EOF Then
                              Me.txt_nombre_articulo = IIf(IsNull(rsaux5!vcha_sub_nombre), "", rsaux5!vcha_sub_nombre)
                           Else
                              Me.txt_nombre_articulo = ""
                           End If
                           rsaux5.Close
                        Else
                           Me.txt_nombre_articulo = ""
                        End If
                     End If
                  Else
                     Me.txt_nombre_articulo = ""
                  End If
                  If Me.txt_nombre_articulo <> "" Then
                     Me.txt_descuento = IIf(IsNull(rsaux10!descuento), 0, rsaux10!descuento)
                     If CDbl(Me.txt_descuento) <= 100 Then
                        If Len(Me.txt_Articulo) = 10 Then
                           For var_j = 1 To lv_clientes.ListItems.Count
                               lv_clientes.ListItems.Item(var_j).Selected = True
                               txt_cliente = Me.lv_clientes.selectedItem
                               If lv_clientes.selectedItem.SubItems(2) = "*" Then
                                                                                                                                                         
                                  VAR_CODIGO_AUX = Me.txt_Articulo
                                  For var_i = 0 To 9
                                      var_codigo = Trim(VAR_CODIGO_AUX) + Trim(CStr(var_i))
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
                                      Me.txt_Articulo = var_codigo
                                      rs.Open "SElECT * from TB_DESCUENTOS_PROMOCION_CLIENTES where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id = '" + Me.txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                                      If Not rs.EOF Then
                                         If Me.chk_marca = 0 Then
                                            var_marca = ""
                                         Else
                                            var_marca = "*"
                                         End If
                                         rsaux.Open "update TB_DESCUENTOS_PROMOCION_CLIENTES set floa_dpr_descuento = " + txt_descuento + ", char_dpr_marca = '" + var_marca + "', DTIM_DPR_FECHA_INICIO = " + var_fecha_inicio + ", DTIM_DPR_FECHA_FIN = " + var_fecha_fin + " + 1 -.00001, VCHA_DPR_USUARIO = '" + var_clave_usuario_global + "',VCHA_DPR_MAQUINA='" + fun_NombrePc + "', DTIM_DPR_FECHA = GETDATE()  where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id = '" + Me.txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                                      Else
                                         If Me.chk_marca = 0 Then
                                            var_marca = ""
                                         Else
                                            var_marca = "*"
                                         End If
                                         rsaux.Open "insert into TB_DESCUENTOS_PROMOCION_CLIENTES (vcha_cli_clave_id, vcha_art_articulo_id, floa_dpr_descuento, char_dpr_marca, DTIM_DPR_FECHA_INICIO, DTIM_DPR_FECHA_FIN, VCHA_DPR_USUARIO, VCHA_DPR_MAQUINA, DTIM_DPR_FECHA) values ('" + txt_cliente + "', '" + Me.txt_Articulo + "', " + Me.txt_descuento + ", '" + var_marca + "'," + var_fecha_inicio + "," + var_fecha_fin + "+1-.00001,'" + var_clave_usuario_global + "','" + fun_NombrePc + "',GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                                      End If
                                      rs.Close
                                  Next var_i
                                  Me.txt_Articulo = VAR_CODIGO_AUX
                               End If
                           Next var_j
                        End If
                        
                        
                        
                        
                        If Len(Trim(Me.txt_Articulo)) = 5 Then
                           If IsNumeric(Me.txt_descuento) Then
                              VAR_CODIGO_AUX = Me.txt_Articulo
                              For var_j = 1 To lv_clientes.ListItems.Count
                                  lv_clientes.ListItems.Item(var_j).Selected = True
                                  txt_cliente = Me.lv_clientes.selectedItem
                                  If lv_clientes.selectedItem.SubItems(2) = "*" Then
                                     If rsaux5.State = 1 Then
                                        rsaux5.Close
                                     End If
                                     rsaux5.Open "select distinct substring(vcha_art_articulo_id,1,10) as codigo from tb_articulos where vcha_Art_articulo_id like '" + Me.txt_Articulo + "%'", cnn, adOpenDynamic, adLockOptimistic
                                     While Not rsaux5.EOF
                                           Me.txt_Articulo = rsaux5!codigo
                                           For var_i = 0 To 9
                                               var_codigo = Trim(rsaux5!codigo) + Trim(CStr(var_i))
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
                                               Me.txt_Articulo = var_codigo
                                               rs.Open "sElect * from TB_DESCUENTOS_PROMOCION_CLIENTES where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id = '" + Me.txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                                               If Not rs.EOF Then
                                                  If Me.chk_marca = 0 Then
                                                     var_marca = ""
                                                  Else
                                                     var_marca = "*"
                                                  End If
                                                  rsaux.Open "update TB_DESCUENTOS_PROMOCION_CLIENTES set floa_dpr_descuento = " + Me.txt_descuento + ", char_dpr_marca = '" + var_marca + "', DTIM_DPR_FECHA_INICIO = " + var_fecha_inicio + ", DTIM_DPR_FECHA_FIN = " + var_fecha_fin + "+1-.00001, vcha_dpr_usuario = '" + var_clave_usuario_global + "', vcha_dpr_maquina = '" + fun_NombrePc + "', dtim_dpr_fecha = getdate() where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id = '" + Me.txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                                               Else
                                                  If Me.chk_marca = 0 Then
                                                     var_marca = ""
                                                  Else
                                                     var_marca = "*"
                                                  End If
                                                  rsaux.Open "insert into TB_DESCUENTOS_PROMOCION_CLIENTES (vcha_cli_clave_id, vcha_art_articulo_id, floa_dpr_descuento, char_dpr_marca, DTIM_DPR_FECHA_INICIO, DTIM_DPR_FECHA_FIN,vcha_dpr_usuario, vcha_dpr_maquina, dtim_dpr_fecha) values ('" + txt_cliente + "', '" + Me.txt_Articulo + "', " + Me.txt_descuento + ", '" + var_marca + "'," + var_fecha_inicio + "," + var_fecha_fin + "+1-.000001,'" + var_clave_usuario_global + "','" + fun_NombrePc + "',getdate())", cnn, adOpenDynamic, adLockOptimistic
                                               End If
                                               rs.Close
                                           Next var_i
                                           rsaux5.MoveNext
                                     Wend
                                  End If
                                  Me.txt_Articulo = VAR_CODIGO_AUX
                              Next var_j
                           End If
                        End If
                           
                           
                     End If
                  End If
                  rsaux10.MoveNext
            Wend
            rsaux10.Close
            MsgBox "Se a terminado de actualizar los descuentos", vbOKOnly, "ATENCION"

            
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
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
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
   
   
   Exit Sub
salir:
   MsgBox "A surgido un error al leer el archivo ofertas.xls", vbOKOnly, "ATENCION"
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
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
End Sub

Private Sub cmd_eliminar_Click()
   For var_j = 1 To lv_clientes.ListItems.Count
       lv_clientes.ListItems.Item(var_j).Selected = True
       If lv_clientes.selectedItem.SubItems(2) = "*" Then
          txt_cliente = lv_clientes.selectedItem
          If Trim(txt_cliente) <> "" Then
             If Trim(Me.txt_Articulo) <> "" Then
                var_si = MsgBox("¿Desea eliminar el registro?", vbYesNo, "ATENCION")
                If var_si = 6 Then
                   rs.Open "DELETE TB_DESCUENTOS_PROMOCION_CLIENTES where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id like '" + Me.txt_Articulo + "%'", cnn, adOpenDynamic, adLockOptimistic
                End If
             Else
             End If
          End If
       End If
   Next var_j
End Sub

Private Sub cmd_guardar_Click()
   Dim var_marca As String
   Dim VERIFICADOR As Integer
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            var_dia = Day(Me.txt_inicio)
            var_mes = Month(Me.txt_inicio)
            var_año = Year(Me.txt_inicio)
            If Len(CStr(var_dia)) = 1 Then
               var_dia_str = "0" + CStr(var_dia)
            Else
               var_dia_str = CStr(var_dia)
            End If
            If Len(CStr(var_mes)) = 1 Then
               var_mes_str = "0" + CStr(var_mes)
            Else
               var_mes_str = CStr(var_mes)
            End If
            If Len(CStr(var_año)) = 1 Then
               var_año_str = "200" + CStr(var_año)
            Else
               If Len(CStr(var_año)) = 2 Then
                  var_año_str = "20" + CStr(var_año)
               Else
                  If Len(CStr(var_año)) = 3 Then
                     var_año_str = "2" + CStr(var_año)
                  Else
                     var_año_str = CStr(var_año)
                  End If
               End If
            End If
            var_fecha_inicio = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
            
            var_dia = Day(Me.txt_fin)
            var_mes = Month(Me.txt_fin)
            var_año = Year(Me.txt_fin)
            If Len(CStr(var_dia)) = 1 Then
               var_dia_str = "0" + CStr(var_dia)
            Else
               var_dia_str = CStr(var_dia)
            End If
            If Len(CStr(var_mes)) = 1 Then
               var_mes_str = "0" + CStr(var_mes)
            Else
               var_mes_str = CStr(var_mes)
            End If
            If Len(CStr(var_año)) = 1 Then
               var_año_str = "200" + CStr(var_año)
            Else
               If Len(CStr(var_año)) = 2 Then
                  var_año_str = "20" + CStr(var_año)
               Else
                  If Len(CStr(var_año)) = 3 Then
                     var_año_str = "2" + CStr(var_año)
                  Else
                     var_año_str = CStr(var_año)
                  End If
               End If
            End If
            var_fecha_fin = "{d '" + var_año_str + "-" + var_mes_str + "-" + var_dia_str + "'}"
            
            
            If Len(Trim(Me.txt_Articulo)) = 10 Then
               If Trim(Me.txt_Articulo) <> "" Then
                  If IsNumeric(Me.txt_descuento) Then
                     If CDbl(Me.txt_descuento) <= 100 Then
                        VAR_CODIGO_AUX = Me.txt_Articulo
                        For var_j = 1 To lv_clientes.ListItems.Count
                            lv_clientes.ListItems.Item(var_j).Selected = True
                            txt_cliente = Me.lv_clientes.selectedItem
                            If lv_clientes.selectedItem.SubItems(2) = "*" Then
                               For var_i = 0 To 9
                                   var_codigo = Trim(VAR_CODIGO_AUX) + Trim(CStr(var_i))
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
                                   Me.txt_Articulo = var_codigo
                                   rs.Open "SElECT * from TB_DESCUENTOS_PROMOCION_CLIENTES where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id = '" + Me.txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                                   If Not rs.EOF Then
                                      If Me.chk_marca = 0 Then
                                         var_marca = ""
                                      Else
                                         var_marca = "*"
                                      End If
                                      rsaux.Open "update TB_DESCUENTOS_PROMOCION_CLIENTES set floa_dpr_descuento = " + txt_descuento + ", char_dpr_marca = '" + var_marca + "', DTIM_DPR_FECHA_INICIO = " + var_fecha_inicio + ", DTIM_DPR_FECHA_FIN = " + var_fecha_fin + " + 1 -.00001, VCHA_DPR_USUARIO = '" + var_clave_usuario_global + "',VCHA_DPR_MAQUINA='" + fun_NombrePc + "', DTIM_DPR_FECHA = GETDATE()  where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id = '" + Me.txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                                   Else
                                      If Me.chk_marca = 0 Then
                                         var_marca = ""
                                      Else
                                         var_marca = "*"
                                      End If
                                      rsaux.Open "insert into TB_DESCUENTOS_PROMOCION_CLIENTES (vcha_cli_clave_id, vcha_art_articulo_id, floa_dpr_descuento, char_dpr_marca, DTIM_DPR_FECHA_INICIO, DTIM_DPR_FECHA_FIN, VCHA_DPR_USUARIO, VCHA_DPR_MAQUINA, DTIM_DPR_FECHA) values ('" + txt_cliente + "', '" + Me.txt_Articulo + "', " + Me.txt_descuento + ", '" + var_marca + "'," + var_fecha_inicio + "," + var_fecha_fin + "+1-.00001,'" + var_clave_usuario_global + "','" + fun_NombrePc + "',GETDATE())", cnn, adOpenDynamic, adLockOptimistic
                                   End If
                                   rs.Close
                               Next var_i
                            End If
                            Me.txt_Articulo = VAR_CODIGO_AUX
                        Next var_j
                        MsgBox "Se a terminado de actualizar los descuentos", vbOKOnly, "ATENCION"
                     Else
                        MsgBox "El descuento no debe de ser mayo al 100%", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
                  End If
               Else
                  MsgBox "Se debe de indicar un artículo", vbOKOnly, "ATENCION"
               End If
            Else
               If Len(Trim(Me.txt_Articulo)) = 5 Then
                  If Trim(Me.txt_Articulo) <> "" Then
                     If IsNumeric(Me.txt_descuento) Then
                        If CDbl(Me.txt_descuento) <= 100 Then
                           VAR_CODIGO_AUX = Me.txt_Articulo
                           For var_j = 1 To lv_clientes.ListItems.Count
                               lv_clientes.ListItems.Item(var_j).Selected = True
                               txt_cliente = Me.lv_clientes.selectedItem
                               If lv_clientes.selectedItem.SubItems(2) = "*" Then
                                  If rsaux5.State = 1 Then
                                     rsaux5.Close
                                  End If
                                  rsaux5.Open "select distinct substring(vcha_art_articulo_id,1,10) as codigo from tb_articulos where vcha_Art_articulo_id like '" + Me.txt_Articulo + "%'", cnn, adOpenDynamic, adLockOptimistic
                                  While Not rsaux5.EOF
                                        Me.txt_Articulo = rsaux5!codigo
                                        For var_i = 0 To 9
                                            var_codigo = Trim(rsaux5!codigo) + Trim(CStr(var_i))
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
                                            Me.txt_Articulo = var_codigo
                                            rs.Open "sElect * from TB_DESCUENTOS_PROMOCION_CLIENTES where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id = '" + Me.txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                                            If Not rs.EOF Then
                                               If Me.chk_marca = 0 Then
                                                  var_marca = ""
                                               Else
                                                  var_marca = "*"
                                               End If
                                               rsaux.Open "update TB_DESCUENTOS_PROMOCION_CLIENTES set floa_dpr_descuento = " + Me.txt_descuento + ", char_dpr_marca = '" + var_marca + "', DTIM_DPR_FECHA_INICIO = " + var_fecha_inicio + ", DTIM_DPR_FECHA_FIN = " + var_fecha_fin + "+1-.00001, vcha_dpr_usuario = '" + var_clave_usuario_global + "', vcha_dpr_maquina = '" + fun_NombrePc + "', dtim_dpr_fecha = getdate() where vcha_cli_clave_id = '" + txt_cliente + "' and vcha_Art_articulo_id = '" + Me.txt_Articulo + "'", cnn, adOpenDynamic, adLockOptimistic
                                            Else
                                               If Me.chk_marca = 0 Then
                                                  var_marca = ""
                                               Else
                                                  var_marca = "*"
                                               End If
                                               rsaux.Open "insert into TB_DESCUENTOS_PROMOCION_CLIENTES (vcha_cli_clave_id, vcha_art_articulo_id, floa_dpr_descuento, char_dpr_marca, DTIM_DPR_FECHA_INICIO, DTIM_DPR_FECHA_FIN,vcha_dpr_usuario, vcha_dpr_maquina, dtim_dpr_fecha) values ('" + txt_cliente + "', '" + Me.txt_Articulo + "', " + Me.txt_descuento + ", '" + var_marca + "'," + var_fecha_inicio + "," + var_fecha_fin + "+1-.000001,'" + var_clave_usuario_global + "','" + fun_NombrePc + "',getdate())", cnn, adOpenDynamic, adLockOptimistic
                                            End If
                                            rs.Close
                                        Next var_i
                                        rsaux5.MoveNext
                                  Wend
                               End If
                               Me.txt_Articulo = VAR_CODIGO_AUX
                           Next var_j
                           MsgBox "Se a terminado de actualizar los descuentos", vbOKOnly, "ATENCION"
                        Else
                           MsgBox "El descuento no debe de ser mayo al 100%", vbOKOnly, "ATENCION"
                        End If
                     Else
                        MsgBox "Descuento incorrecto", vbOKOnly, "ATENCION"
                     End If
                  Else
                     MsgBox "Se debe de indicar un artículo", vbOKOnly, "ATENCION"
                  End If
               End If
            End If
         Else
            MsgBox "La fecha inicial debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
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
   If rsaux11.State = 1 Then
      rsaux11.Close
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_Articulo = ""
   Me.txt_nombre_articulo = ""
   Me.txt_descuento = ""
   Me.chk_marca = 0
   n = lv_clientes.ListItems.Count
   For i = 1 To n
      lv_clientes.ListItems.Item(i).Selected = True
      lv_clientes.selectedItem.SubItems(2) = ""
      lv_clientes.ListItems.Item(i).Bold = False
      lv_clientes.ListItems.Item(i).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_clientes.Refresh
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   n = lv_clientes.ListItems.Count
   var_rellena = True
   var_encontro = False
   For i = 1 To n
      lv_clientes.ListItems.Item(i).Selected = True
      If var_encontro = True And lv_clientes.selectedItem.SubItems(2) = "" And var_rellena = True Then
         lv_clientes.selectedItem.SubItems(2) = "*"
         lv_clientes.ListItems.Item(i).Bold = True
         lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      Else
         If var_encontro = True And lv_clientes.selectedItem.SubItems(2) = "*" Then
            var_rellena = False
         End If
      End If
      If lv_clientes.selectedItem.SubItems(2) = "*" And var_encontro = False Then
         var_encontro = True
      End If
   Next i
End Sub

Private Sub Command2_Click()
   i = lv_clientes.selectedItem.Index
   If lv_clientes.selectedItem.SubItems(2) = "*" Then
      lv_clientes.selectedItem.SubItems(2) = ""
      lv_clientes.ListItems.Item(i).Bold = False
      lv_clientes.ListItems.Item(i).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      lv_clientes.Refresh
   Else
      lv_clientes.selectedItem.SubItems(2) = "*"
      lv_clientes.ListItems.Item(i).Bold = True
      lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      lv_clientes.Refresh
   End If
End Sub

Private Sub Command3_Click()
   n = lv_clientes.ListItems.Count
   For i = 1 To n
      lv_clientes.ListItems.Item(i).Selected = True
      If lv_clientes.selectedItem.SubItems(2) = "*" Then
         lv_clientes.selectedItem.SubItems(2) = ""
         lv_clientes.ListItems.Item(i).Bold = False
         lv_clientes.ListItems.Item(i).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
      Else
         lv_clientes.selectedItem.SubItems(2) = "*"
         lv_clientes.ListItems.Item(i).Bold = True
         lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
      End If
   Next i
End Sub

Private Sub Command4_Click()
   n = lv_clientes.ListItems.Count
   For i = 1 To n
      lv_clientes.ListItems.Item(i).Selected = True
      lv_clientes.selectedItem.SubItems(2) = ""
      lv_clientes.ListItems.Item(i).Bold = False
      lv_clientes.ListItems.Item(i).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
      lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
      lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
   Next i
   lv_clientes.Refresh
End Sub

Private Sub Command5_Click()
   n = lv_clientes.ListItems.Count
   For i = 1 To n
      lv_clientes.ListItems.Item(i).Selected = True
      lv_clientes.selectedItem.SubItems(2) = "*"
      lv_clientes.ListItems.Item(i).Bold = True
      lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
      lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
      lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
      lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
      lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
   Next i
   lv_clientes.Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Unload Me
   End If

End Sub

Private Sub Form_Load()
   Top = 300
   Left = 2200
   Me.txt_inicio = Date
   Me.txt_fin = Date
   rs.Open "select DISTINCT VCHA_CLI_CLAVE_ID, VCHA_CLI_NOMBRE from VW_clientes WHERE VCHA_EMP_EMPRESA_ID = '" + var_empresa + "' order by vcha_cli_nombre", cnn, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         Set list_item = lv_clientes.ListItems.Add(, , rs!vcha_cli_clave_id)
         list_item.SubItems(1) = IIf(IsNull(rs!VCHA_CLI_NOMBRE), "", rs!VCHA_CLI_NOMBRE)
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub



Private Sub lv_clientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_clientes, ColumnHeader)
End Sub

Private Sub lv_clientes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      i = lv_clientes.selectedItem.Index
      If lv_clientes.selectedItem.SubItems(2) = "*" Then
         lv_clientes.selectedItem.SubItems(2) = ""
         lv_clientes.ListItems.Item(i).Bold = False
         lv_clientes.ListItems.Item(i).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = False
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &H80000012
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &H80000012
         lv_clientes.Refresh
      Else
         lv_clientes.selectedItem.SubItems(2) = "*"
         lv_clientes.ListItems.Item(i).Bold = True
         lv_clientes.ListItems.Item(i).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(1).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(2).Bold = True
         lv_clientes.ListItems.Item(i).ListSubItems(1).ForeColor = &HFF0000
         lv_clientes.ListItems.Item(i).ListSubItems(2).ForeColor = &HFF0000
         lv_clientes.Refresh
      End If
   End If
End Sub

Private Sub txt_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_articulo_LostFocus()
   
   If Trim(Me.txt_Articulo) <> "" Then
      If Len(Me.txt_Articulo) = 10 Then
         rs.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_ART_ARTICULO_ID LIKE  '" + Me.txt_Articulo + "%'", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_nombre_articulo = IIf(IsNull(rs!VCHA_ART_NOMBRE_ESPAÑOL), "", rs!VCHA_ART_NOMBRE_ESPAÑOL)
         Else
            Me.txt_Articulo = ""
            Me.txt_nombre_articulo = ""
         End If
         rs.Close
      Else
         If Len(Trim(Me.txt_Articulo)) = 5 Then
            If rsaux5.State = 1 Then
               rsaux5.Close
            End If
            rsaux5.Open "select * from tb_subdivisiones where vcha_tpr_tipo_producto_id = '" + Mid(Me.txt_Articulo, 1, 1) + "' and vcha_div_division_id = '" + Mid(Me.txt_Articulo, 2, 2) + "' and vcha_sub_subdivision_id = '" + Mid(Me.txt_Articulo, 4, 2) + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux5.EOF Then
               Me.txt_nombre_articulo = IIf(IsNull(rsaux5!vcha_sub_nombre), "", rsaux5!vcha_sub_nombre)
            Else
               Me.txt_nombre_articulo = ""
            End If
            rsaux5.Close
         Else
            MsgBox "Clave de artículo incorrecta", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      Me.txt_nombre_articulo = ""
   End If
End Sub




Private Sub txt_descuento_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes.Value = CDate(Me.txt_fin)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes.Value = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes.Value = Date
      End If
      frmcalendario.Show 1
      txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_nombre_articulo_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 And KeyAscii <> 27 Then
      KeyAscii = 0
   End If
   Call pro_enfoque(KeyAscii)
End Sub


