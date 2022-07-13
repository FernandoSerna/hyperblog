VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_rendimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rendimiento"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_eliminar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      Picture         =   "frmoracle_rendimiento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_periodo 
      Height          =   1080
      Left            =   1560
      TabIndex        =   27
      Top             =   240
      Width           =   4245
      Begin VB.CommandButton cmd_reporte 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   0
         Picture         =   "frmoracle_rendimiento.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Imprimir Alt + I"
         Top             =   120
         Width           =   330
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   540
         Width           =   1140
      End
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   555
         Width           =   1140
      End
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   0
         TabIndex        =   33
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2325
         TabIndex        =   31
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   615
         Width           =   420
      End
   End
   Begin VB.Frame frm_buscar 
      Height          =   735
      Left            =   1080
      TabIndex        =   25
      Top             =   360
      Width           =   2175
      Begin VB.TextBox txt_buscar 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Picture         =   "frmoracle_rendimiento.frx":0634
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Buscar Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   480
      TabIndex        =   21
      Top             =   360
      Width           =   5685
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   45
         TabIndex        =   22
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Volumen"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000FF&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   5610
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmoracle_rendimiento.frx":0736
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      Picture         =   "frmoracle_rendimiento.frx":0838
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      Picture         =   "frmoracle_rendimiento.frx":093A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6240
      Picture         =   "frmoracle_rendimiento.frx":0A3C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   6495
      Begin VB.TextBox txt_folio 
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_rendimiento 
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txt_nombre_unidad 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txt_unidad 
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt_nombre_ruta 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txt_ruta 
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txt_nombre_chofer 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txt_chofer 
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt_fecha 
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         Height          =   195
         Left            =   4680
         TabIndex        =   19
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Rendimiento:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Ruta:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Chofer:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmoracle_rendimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_ventana As Integer

Private Sub cmd_buscar_Click()
   Me.frm_periodo.Visible = False
   Me.frm_buscar.Visible = True
   Me.txt_buscar = ""
   Me.txt_buscar.SetFocus
End Sub

Private Sub cmd_eliminar_Click()
   If Me.txt_folio <> "" Then
      var_si = MsgBox("Desea eliminar el registro?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         rs.Open "DELETE FROM TB_ORACLE_rENDIMIENTO WHERE FOLIO = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
         MsgBox "Se a eliminado el registro", vbOKOnly, "ATENCION"
         Me.txt_fecha = ""
         Me.txt_chofer = ""
         Me.txt_nombre_chofer = ""
         Me.txt_ruta = ""
         Me.txt_nombre_ruta = ""
         Me.txt_unidad = ""
         Me.txt_nombre_unidad = ""
         Me.txt_rendimiento = ""
         Me.txt_folio = ""
      End If
   End If
End Sub

Private Sub cmd_guardar_Click()
   Me.frm_periodo.Visible = False
    If IsDate(Me.txt_fecha) Then
       If Me.txt_chofer <> "" Then
          If Me.txt_ruta <> "" Then
             If Me.txt_unidad <> "" Then
                If IsNumeric(Me.txt_rendimiento) Then
                   var_dia = CStr(Day(CDate(Me.txt_fecha)))
                   var_mes = CStr(Month(CDate(Me.txt_fecha)))
                   var_año = CStr(Year(CDate(Me.txt_fecha)))
                   If Len(Trim(var_dia)) = 1 Then
                      var_dia = "0" + var_dia
                   End If
                   If Len(Trim(var_mes)) = 1 Then
                      var_mes = "0" + var_mes
                   End If
            
                   var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                   
                   If IsNumeric(Me.txt_folio) Then
                      rs.Open "UPDATE TB_ORACLE_RENDIMIENTO SET CHOFER = '" + Me.txt_chofer + "', RUTA = '" + Me.txt_ruta + "', UNIDAD = '" + Me.txt_unidad + "', RENDIMIENTO = " + Me.txt_rendimiento + ", fecha = " + var_fecha + " WHERE FOLIO = " + Me.txt_folio, cnn, adOpenDynamic, adLockOptimistic
                      MsgBox "Se a actualizado la información", vbOKOnly, "ATENCION"
                   Else
                      rs.Open "INSERT INTO TB_ORACLE_RENDIMIENTO (CHOFER, RUTA, UNIDAD, RENDIMIENTO, fecha, USUARIO) VALUES ('" + Me.txt_chofer + "','" + Me.txt_ruta + "','" + Me.txt_unidad + "'," + Me.txt_rendimiento + "," + var_fecha + ",'" + var_clave_usuario_global + "')", cnn, adOpenDynamic, adLockOptimistic
                      rs.Open "SELECT MAX(FOLIO) FROM TB_ORACLE_rENDIMIENTO", cnn, adOpenDynamic, adLockOptimistic
                      MsgBox "Se a insertado el registro", vbOKOnly, "ATENCION"
                      If Not rs.EOF Then
                         Me.txt_folio = rs(0).Value
                      End If
                      rs.Close
                   End If
                Else
                   MsgBox "Rendimiento incorrecto", vbOKOnly, "ATENCION"
                End If
             Else
                MsgBox "Unidad incorrecta", vbOKOnly, "ATENCION"
             End If
          Else
             MsgBox "Ruta incorrecta", vbOKOnly, "ATENCION"
          End If
       Else
          MsgBox "Chofer incorrecto", vbOKOnly, "ATENCION"
       End If
    Else
       MsgBox "Fecha incorrecta", vbOKOnly, "ATENCION"
    End If
    
End Sub

Private Sub cmd_imprimir_Click()
   Me.frm_periodo.Visible = False
   Me.frm_periodo.Visible = True
   Me.txt_inicio = Date
   Me.txt_fin = Date
   Me.txt_inicio.SetFocus
End Sub

Private Sub cmd_nuevo_Click()
   Me.frm_periodo.Visible = False
   Me.txt_fecha = Date
   Me.txt_chofer = ""
   Me.txt_nombre_chofer = ""
   Me.txt_ruta = ""
   Me.txt_nombre_ruta = ""
   Me.txt_unidad = ""
   Me.txt_nombre_unidad = ""
   Me.txt_rendimiento = ""
   Me.txt_folio = ""
   Me.txt_fecha.SetFocus
End Sub

Private Sub cmd_reporte_Click()
   If IsDate(Me.txt_inicio) Then
      If IsDate(Me.txt_fin) Then
         If CDate(Me.txt_inicio) <= CDate(Me.txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(inte_Tem_Consecutivo) from tb_Temp_oracle_Rendimiento", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
            Else
               var_consecutivo = 1
            End If
            rs.Close
            rs.Open "insert into tb_temp_oracle_rendimiento (inte_Tem_Consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_dia = CStr(Day(CDate(Me.txt_inicio)))
            var_mes = CStr(Month(CDate(Me.txt_inicio)))
            var_año = CStr(Year(CDate(Me.txt_inicio)))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_inicio = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_dia = CStr(Day(CDate(Me.txt_fin) + 1))
            var_mes = CStr(Month(CDate(Me.txt_fin) + 1))
            var_año = CStr(Year(CDate(Me.txt_fin) + 1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            var_cadena = "SELECT dbo.TB_ORACLE_RENDIMIENTO.USUARIO,dbo.TB_ORACLE_RENDIMIENTO.REGISTRO, dbo.TB_ORACLE_RENDIMIENTO.FOLIO, dbo.TB_ORACLE_RENDIMIENTO.FECHA, dbo.TB_ORACLE_RENDIMIENTO.CHOFER, dbo.TB_ORACLE_RENDIMIENTO.RUTA, dbo.TB_ORACLE_RENDIMIENTO.UNIDAD, dbo.TB_ORACLE_RENDIMIENTO.RENDIMIENTO , dbo.TB_CHOFERES.VCHA_CHO_NOMBRE, dbo.TB_ORACLE_TRANSPORTES.NOMBRE FROM dbo.TB_ORACLE_RENDIMIENTO INNER JOIN dbo.TB_CHOFERES ON dbo.TB_ORACLE_RENDIMIENTO.CHOFER = dbo.TB_CHOFERES.VCHA_CHO_CHOFER_ID INNER JOIN"
            var_cadena = var_cadena + " dbo.TB_ORACLE_TRANSPORTES ON dbo.TB_ORACLE_RENDIMIENTO.UNIDAD = dbo.TB_ORACLE_TRANSPORTES.CLAVE WHERE FECHA >= " + var_fecha_inicio + " and fecha < " + var_fecha_fin
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  var_dia = CStr(Day(CDate(rs!Fecha)))
                  var_mes = CStr(Month(CDate(rs!Fecha) + 1))
                  var_año = CStr(Year(CDate(rs!Fecha) + 1))
                  If Len(Trim(var_dia)) = 1 Then
                     var_dia = "0" + var_dia
                  End If
                  If Len(Trim(var_mes)) = 1 Then
                     var_mes = "0" + var_mes
                  End If
                  var_fecha = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
                  rsaux1.Open "SELECT * FROM TB_USUARIOS WHERE VCHA_USU_USUARIO_ID = '" + IIf(IsNull(rs!USUARIO), "", rs!USUARIO) + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     var_nombre_usuario_RENDIMIENTO = IIf(IsNull(rsaux1!VCHA_USU_NOMBRE), "", rsaux1!VCHA_USU_NOMBRE) + " " + IIf(IsNull(rsaux1!VCHA_USU_APELLIDOS), "", rsaux1!VCHA_USU_APELLIDOS)
                  Else
                     var_nombre_usuario_RENDIMIENTO = ""
                  End If
                  rsaux1.Close
                  var_cadena = "insert into tb_temp_oracle_rendimiento (inte_Tem_consecutivo, fecha_inicio, fecha_fin, fecha, chofer, nombre_chofer, ruta, nombre_ruta, unidad, nombre_unidad, rendimiento, USUARIO, FECHA_HORA_REGISTRO)"
                  var_cadena = var_cadena + "           values (" + CStr(var_consecutivo) + "," + var_fecha_inicio + "," + var_fecha_fin + "-1," + var_fecha + ",'" + rs!chofer + "','" + rs!vcha_cho_nombre + "','" + rs!ruta + "','','" + rs!unidad + "','" + rs!nombre + "'," + CStr(rs!rendimiento) + ",'" + Mid(var_nombre_usuario_RENDIMIENTO, 1, 50) + "','" + CStr(IIf(IsNull(rs!REGISTRO), "", rs!REGISTRO)) + "')"
                  rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select distinct ruta from tb_Temp_oracle_rendimiento where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux1.Open "select * from xxvia_Tb_rutas_distribucion where ruta = '" + IIf(IsNull(rs!ruta), "", rs!ruta) + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
                  If Not rsaux1.EOF Then
                     rsaux2.Open "update tb_Temp_oracle_rendimiento set nombre_ruta = '" + rsaux1!nombre_ruta + "' where ruta = '" + rs!ruta + "' and inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                  End If
                  rsaux1.Close
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select unidad, sum(rendimiento)/count(*) media  from tb_temp_oracle_rendimiento where inte_tem_consecutivo = " + CStr(var_consecutivo) + " and unidad is not null group by unidad ", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux1.Open "update tb_Temp_oracle_rendimiento set media = " + CStr(rs!media) + " where unidad = '" + rs!unidad + "'", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
            rs.Close
            rs.Open "select FECHA_INICIO, FECHA_FIN, FECHA, UNIDAD, NOMBRE_UNIDAD, RUTA, NOMBRE_RUTA, CHOFER, NOMBRE_CHOFER, RENDIMIENTO, MEDIA, USUARIO, FECHA_HORA_REGISTRO from tb_Temp_oracle_rendimiento WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND UNIDAD IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
            Dim iFila As Long, iCol As Integer, i As Integer
            Set oexcel = CreateObject("Excel.Application")
            Set oWBook = oexcel.Workbooks.Add
            Set oSheet = oWBook.Worksheets(1)
            oSheet.Name = "RENDIMIENTO"
            Screen.MousePointer = vbHourglass
            iFila = 1
            iFila2 = 1
            iCol2 = 1
            iCol = 1
            'MsgBox var_cadena
            For i = 0 To rs.Fields.Count - 1
                oSheet.Cells(iFila, i + 1) = rs.Fields(i).Name
                oSheet.Cells(iFila, i + 1).Font.Bold = True
            Next
            iFila = iFila + 1
            
            With oSheet
                 ' carga los registros del recordset
                 .Cells(iFila, iCol).CopyFromRecordset rs
                 'oExcel.Columns(1).Select
                 'oExcel.Selection.NumberFormat = "#,##0.00"
                 'oExcel.Columns(1).Select
                 'oExcel.Selection.Font.Color = vbRed
                 .Columns.AutoFit ' ajusta el ancho de las columnas
            End With
            oWBook.SaveAs "c:\reportessid\rendemiento_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
            oexcel.Visible = True
            Set oexcel = Nothing
            Screen.MousePointer = vbDefault
            
            rs.Close
            rs.Open "delete from tb_temp_oracle_Rendimiento where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
            Me.frm_periodo.Visible = False
         Else
            MsgBox "La fecha de inicio no debe de ser mayor a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha inicial incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Top = 2000
    Left = 2000
    Me.frm_lista.Visible = False
    Me.frm_buscar.Visible = False
    Me.frm_periodo.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If var_ventana = 1 Then
         Me.txt_chofer = Me.lv_lista.selectedItem
         Me.txt_nombre_chofer = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_chofer.SetFocus
      End If
      If var_ventana = 2 Then
         Me.txt_ruta = Me.lv_lista.selectedItem
         Me.txt_nombre_ruta = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_ruta.SetFocus
      End If
      If var_ventana = 3 Then
         Me.txt_unidad = Me.lv_lista.selectedItem
         Me.txt_nombre_unidad = Me.lv_lista.selectedItem.SubItems(1)
         Me.txt_unidad.SetFocus
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frm_buscar.Visible = False
   End If
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_buscar) Then
         rs.Open "select * from tb_oracle_rendimiento where folio = " + Me.txt_buscar, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_fecha = rs!Fecha
            rsaux1.Open "select * from tb_choferes where vcha_cho_chofer_id = '" + rs!chofer + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_chofer = rsaux1!vcha_cho_chofer_id
               Me.txt_nombre_chofer = rsaux1!vcha_cho_nombre
            Else
               Me.txt_chofer = ""
               Me.txt_nombre_chofer = ""
            End If
            rsaux1.Close
            rsaux1.Open "select * from xxvia_tb_rutas_distribucion where ruta = '" + rs!ruta + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_ruta = rsaux1!ruta
               Me.txt_nombre_ruta = rsaux1!nombre_ruta
            Else
               Me.txt_ruta = ""
               Me.txt_nombre_ruta = ""
            End If
            rsaux1.Close
            rsaux1.Open "select * from tb_oracle_transportes where clave = '" + rs!unidad + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux1.EOF Then
               Me.txt_unidad = rsaux1!clave
               Me.txt_nombre_unidad = rsaux1!nombre
            Else
               Me.txt_unidad = ""
               Me.txt_nombre_unidad = ""
            End If
            rsaux1.Close
            Me.txt_folio = Me.txt_buscar
            Me.txt_rendimiento = rs!rendimiento
         Else
            MsgBox "El folio no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
         Me.txt_chofer.SetFocus
      Else
         MsgBox "Folio incorrecto", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub txt_buscar_LostFocus()
   Me.frm_buscar.Visible = False
End Sub

Private Sub txt_chofer_Change()
   Me.txt_nombre_chofer = ""
End Sub

Private Sub txt_chofer_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_chofer_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 1
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_choferes ", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!vcha_cho_chofer_id)
            list_item.SubItems(1) = IIf(IsNull(rs!vcha_cho_nombre), "", rs!vcha_cho_nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "CHOFERES"
      VAR_TIPO_LISTA = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 2800
         lv_lista.ColumnHeaders(3).Width = 1400
      Else
         lv_lista.ColumnHeaders(2).Width = 3000.18
         lv_lista.ColumnHeaders(3).Width = 1400
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_chofer_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_chofer_LostFocus()
    If Me.txt_chofer <> "" Then
       rs.Open "select * from tb_choferes where vcha_cho_chofer_id = '" + Me.txt_chofer + "'", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          Me.txt_nombre_chofer = IIf(IsNull(rs!vcha_cho_nombre), "", rs!vcha_cho_nombre)
       Else
          MsgBox "Clave incorrecta", vbOKOnly, "ATENCION"
          Me.txt_chofer = ""
          Me.txt_nombre_chofer = ""
       End If
       rs.Close
    End If
End Sub

Private Sub txt_fecha_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_fecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fecha) Then
         frmcalendario.mes = CDate(Me.txt_fecha)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fecha = var_fecha_general
   End If
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_fin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_fin) Then
         frmcalendario.mes = CDate(Me.txt_fin)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_fin = var_fecha_general
   End If
End Sub

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       Me.frm_periodo.Visible = False
    End If
    If KeyAscii = 13 Then
       Me.cmd_reporte.SetFocus
    End If
End Sub

Private Sub txt_folio_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_inicio_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      If IsDate(Me.txt_inicio) Then
         frmcalendario.mes = CDate(Me.txt_inicio)
      Else
         frmcalendario.mes = Date
      End If
      frmcalendario.Show 1
      Me.txt_inicio = var_fecha_general
   End If
End Sub

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       Me.frm_periodo.Visible = False
    End If
    If KeyAscii = 13 Then
       Me.txt_fin.SetFocus
    End If
End Sub

Private Sub txt_nombre_chofer_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_nombre_chofer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_ruta_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_nombre_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_nombre_unidad_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_nombre_unidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call pro_enfoque(KeyAscii)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_rendimiento_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_rendimiento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_guardar.SetFocus
   End If
End Sub

Private Sub txt_ruta_Change()
   Me.txt_nombre_ruta = ""
End Sub

Private Sub txt_ruta_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_ruta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 2
      lv_lista.ListItems.Clear
      rs.Open "select * from xxvia_tb_rutas_distribucion", cnnoracle_4, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!ruta)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "RUTAS"
      VAR_TIPO_LISTA = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 2800
         lv_lista.ColumnHeaders(3).Width = 1400
      Else
         lv_lista.ColumnHeaders(2).Width = 3000.18
         lv_lista.ColumnHeaders(3).Width = 1400
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_ruta_LostFocus()
   If Me.txt_ruta <> "" Then
      rs.Open "SELECT * FROM XXVIA_tB_RUTAS_DISTRIBUCION WHERE RUTA = '" + Me.txt_ruta + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_nombre_ruta = IIf(IsNull(rs!nombre_ruta), "", rs!nombre_ruta)
      Else
         MsgBox "Clave de ruta incorrecta", vbOKOnly, "ATENCION"
         Me.txt_nombre_ruta = ""
         Me.txt_ruta = ""
      End If
      rs.Close
   End If
   
End Sub

Private Sub txt_unidad_Change()
   Me.txt_nombre_unidad = ""
End Sub

Private Sub txt_unidad_GotFocus()
   Me.frm_periodo.Visible = False
End Sub

Private Sub txt_unidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      var_ventana = 3
      lv_lista.ListItems.Clear
      rs.Open "select * from tb_oracle_transportes", cnn, adOpenDynamic, adLockOptimistic
      While Not rs.EOF
            Set list_item = lv_lista.ListItems.Add(, , rs!clave)
            list_item.SubItems(1) = IIf(IsNull(rs!nombre), "", rs!nombre)
            rs.MoveNext
      Wend
      rs.Close
      lbl_lista = "TRANSPORTES"
      VAR_TIPO_LISTA = 100
      Dim var_n As Integer
      var_n = lv_lista.ListItems.Count
      If var_n > 6 Then
         lv_lista.ColumnHeaders(2).Width = 2800
         lv_lista.ColumnHeaders(3).Width = 1400
      Else
         lv_lista.ColumnHeaders(2).Width = 3000.18
         lv_lista.ColumnHeaders(3).Width = 1400
      End If
      frm_lista.Visible = True
      lv_lista.SetFocus
   End If
End Sub

Private Sub txt_unidad_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_unidad_LostFocus()
    If Me.txt_unidad <> "" Then
       rs.Open "SELECT * FROM tb_oracle_transportes WHERE CLAVE = '" + Me.txt_unidad + "'", cnn, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
          Me.txt_nombre_unidad = IIf(IsNull(rs!nombre), "", rs!nombre)
       Else
          MsgBox "Clave de unidad incorrecta", vbOKOnly, "ATENCION"
          Me.txt_unidad = ""
          Me.txt_nombre_unidad = ""
       End If
       rs.Close
    End If
End Sub
