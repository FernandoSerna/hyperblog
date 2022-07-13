VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnotas_traspasos_plantas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notas de traspasos entre almacenes"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3045
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   7680
      Begin MSComctlLib.ListView lv_notas 
         Height          =   2790
         Left            =   75
         TabIndex        =   1
         Top             =   180
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   4921
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nota"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Planta origen"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Recibidos"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Clave planta origen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nota Origen"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Dias"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Clave planta DESTINO"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmnotas_traspasos_plantas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_clave_unidad_planta As String
Dim var_clave_unidad_origen As String
Public var_str_encabezado_forma As String
Public var_str_nota_envio  As String

Private Sub Form_Load()
   var_nota_traspasos = ""
   Dim strNotaTraspaso As String
   Dim list_item As ListItem
   Dim var_planta_id As String
   'MsgBox var_unidad_organizacional
   rs.Open "select * from tb_plantas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_pla_planta_id = '" + var_planta_transito_global + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
   var_clave_unidad_planta = ""
   If Not rs.EOF Then
      var_clave_unidad_planta = rs!vcha_pla_planta_id
   End If
   Me.Caption = var_str_encabezado_forma
   rs.Close
   'MsgBox var_clave_unidad_planta
   If var_str_encabezado_forma = "" Then
            MsgBox "Favor de asignar valor al encabezado", vbCritical, "SIP"
            Exit Sub
        End If
   Select Case var_clave_movimiento
   
    Case "ETA"
        'rs.Open "select distinct numPlantaOrigen VCHA_TRA_PLANTA_ORIGEN, " & _
        '            "numPlantaDestino  VCHA_TRA_PLANTA_DESTINO, " & _
        '            "VCHA_MOV_MOVIMIENTO_ID, VCHA_TRA_NOTA_ENVIO, " & _
        '            "VCHA_TRA_CALIDAD, MAX(FechaEnvio) DATE_tRA_FECHA_ENVIO, " & _
        '            "VCHA_EMP_EMPRESA_ID, sum(FLOA_TRA_CANTIDAD_ENVIADA ) CANTIDAD_ENVIADA, " & _
        '            "sum(FLOA_TRA_CANTIDAD_RECIBIDA ) CANTIDAD_RECIBIDA  " & _
        '        "from vw_entradaPendientesTransito  " & _
        '        "where numPlantaDestino ='" + var_clave_unidad_planta + "' " & _
        '        "AND VCHA_MOV_MOVIMIENTO_ID IN ( 'SALTRA' ) " & _
        '        "AND FLOA_TRA_CANTIDAD_RECIBIDA = 0 " & _
        '        "GROUP BY numPlantaOrigen, numPlantaDestino,  " & _
        '                "VCHA_MOV_MOVIMIENTO_ID, VCHA_TRA_NOTA_ENVIO, " & _
        '                "VCHA_TRA_CALIDAD, VCHA_EMP_EMPRESA_ID  " & _
        '        "ORDER BY numPlantaOrigen,VCHA_TRA_NOTA_ENVIO  ", _
        'cnn_admcdindustrial, _
        'adOpenDynamic, _
        'adLockOptimistic
            
                
        rs.Open "SELECT VCHA_TRA_PLANTA_ORIGEN, VCHA_TRA_PLANTA_DESTINO, " & _
                        "VCHA_MOV_MOVIMIENTO_ID, VCHA_TRA_NOTA_ENVIO, " & _
                        "VCHA_TRA_CALIDAD, MAX(DATE_TRA_FECHA_ENVIO) AS DATE_tRA_FECHA_ENVIO, " & _
                        "VCHA_EMP_EMPRESA_ID, SUM(CANTIDAD_ENVIADA) AS CANTIDAD_ENVIADA, " & _
                        "SUM(CANTIDAD_RECIBIDA) As CANTIDAD_RECIBIDA " & _
                "From dbo.VW_TRANSITO_ENVIADO_RECIBIDA_TOTALES " & _
                "WHERE (VCHA_TRA_PLANTA_DESTINO = '" + var_clave_unidad_planta + "') " & _
                "AND (CANTIDAD_ENVIADA > CANTIDAD_RECIBIDA) " & _
                "AND (VCHA_MOV_MOVIMIENTO_ID = 'SALTRA') " & _
                "GROUP BY VCHA_TRA_PLANTA_ORIGEN, VCHA_TRA_PLANTA_DESTINO, " & _
                        "VCHA_MOV_MOVIMIENTO_ID, VCHA_TRA_NOTA_ENVIO, " & _
                        "VCHA_TRA_CALIDAD, VCHA_EMP_EMPRESA_ID ", _
            cnn_admcdindustrial, _
            adOpenDynamic, _
            adLockOptimistic
        If rs.RecordCount > 0 Then
            var_planta_id = rs("VCHA_TRA_PLANTA_ORIGEN").Value
        Else
            var_planta_id = ""
        End If
    Case "DPL"
        rs.Open "SELECT VCHA_TRA_PLANTA_ORIGEN, VCHA_TRA_PLANTA_DESTINO, " & _
                        "VCHA_MOV_MOVIMIENTO_ID, VCHA_TRA_NOTA_ENVIO, " & _
                        "VCHA_TRA_CALIDAD, CONVERT( VARCHAR, MAX( DATE_TRA_FECHA_ENVIO),3)  DATE_tRA_FECHA_ENVIO, " & _
                        "VCHA_EMP_EMPRESA_ID, SUM(CANTIDAD_ENVIADA) AS CANTIDAD_ENVIADA, " & _
                        "SUM(CANTIDAD_RECIBIDA) As CANTIDAD_RECIBIDA " & _
                "From dbo.VW_TRANSITO_ENVIADO_RECIBIDA_TOTALES " & _
                "WHERE (VCHA_TRA_PLANTA_origen = '" + var_clave_unidad_planta + "') " & _
                "AND ( CANTIDAD_ENVIADA > CANTIDAD_RECIBIDA ) " & _
                "AND (VCHA_MOV_MOVIMIENTO_ID = 'SALTRA') " & _
                "GROUP BY VCHA_TRA_PLANTA_ORIGEN, VCHA_TRA_PLANTA_DESTINO, " & _
                        "VCHA_MOV_MOVIMIENTO_ID, VCHA_TRA_NOTA_ENVIO, " & _
                        "VCHA_TRA_CALIDAD, VCHA_EMP_EMPRESA_ID ", _
            cnn_admcdindustrial, _
            adOpenDynamic, _
            adLockOptimistic
            lv_notas.ColumnHeaders(2).Text = "Planta Destino"
            If rs.RecordCount > 0 Then
                var_planta_id = rs("VCHA_TRA_PLANTA_ORIGEN").Value
            Else
                var_planta_id = ""
            End If
    Case "EP"
         rs.Open "SELECT ALMACEN_ORIGEN_ID VCHA_TRA_PLANTA_ORIGEN, PLANTA_DESTINO VCHA_TRA_PLANTA_DESTINO, " & _
                        "MOVIMIENTO_ORIGEN VCHA_MOV_MOVIMIENTO_ID, NOTA_ENVIO VCHA_TRA_NOTA_ENVIO, " & _
                        "CALIDAD VCHA_TRA_CALIDAD, TO_char(FECHA_ENVIO,'DD/MM/YYYY')  DATE_tRA_FECHA_ENVIO, " & _
                        "EMPRESA_DESTINO VCHA_EMP_EMPRESA_ID, SUM(CANTIDAD_ENVIADA) AS CANTIDAD_ENVIADA, " & _
                        "SUM(CANTIDAD_RECIBIDA) As CANTIDAD_RECIBIDA, PLANTA_ORIGEN " & _
                "From vw_transito " & _
                "where PLANTA_DESTINO ='" & var_clave_unidad_planta & "' " & _
                " and unidad_destino    ='" & var_unidad_organizacional & "' " & _
                " AND empresa_destino   ='" & var_empresa & "' " & _
                "AND ( CANTIDAD_RECIBIDA = 0 ) " & _
                "grOUP BY ALMACEN_ORIGEN_ID,   PLANTA_DESTINO,   MOVIMIENTO_ORIGEN,  NOTA_ENVIO,  CALIDAD,  TO_char(FECHA_ENVIO,'DD/MM/YYYY'),  EMPRESA_DESTINO, PLANTA_ORIGEN ", _
              cnnoracle, _
            adOpenDynamic, _
            adLockOptimistic
    Case Else
        rs.Open "SELECT VCHA_TRA_PLANTA_ORIGEN, VCHA_TRA_PLANTA_DESTINO, VCHA_MOV_MOVIMIENTO_ID, CONVERT( VARCHAR ,DATE_TRA_FECHA_ENVIO,6)  DATE_tRA_FECHA_ENVIO, VCHA_TRA_CALIDAD, DATE_tRA_FECHA_ENVIO, VCHA_EMP_EMPRESA_ID, SUM(CANTIDAD_ENVIADA) AS CANTIDAD_ENVIADA, SUM(CANTIDAD_RECIBIDA) As CANTIDAD_RECIBIDA From dbo.VW_TRANSITO_ENVIADO_RECIBIDA_TOTALES WHERE (VCHA_TRA_PLANTA_DESTINO = '" + var_clave_unidad_planta + "') AND (CANTIDAD_ENVIADA > CANTIDAD_RECIBIDA) AND (VCHA_MOV_MOVIMIENTO_ID = '" + var_clave_movimiento + "') GROUP BY VCHA_TRA_PLANTA_ORIGEN, VCHA_TRA_PLANTA_DESTINO, VCHA_MOV_MOVIMIENTO_ID, VCHA_TRA_NOTA_ENVIO, VCHA_TRA_CALIDAD, VCHA_EMP_EMPRESA_ID ", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
        var_planta_id = rs("VCHA_TRA_PLANTA_ORIGEN").Value
        'rs.Open "select * from VW_TRANSITO_ENVIADO_RECIBIDA_TOTALES where VCHA_TRA_PLANTA_DESTINO = '" + var_clave_unidad_planta + "' and cantidad_enviada > cantidad_recibida and vcha_mov_movimiento_id = '" + var_clave_movimiento + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
    End Select
     
   numero_items_licencias = 0
   While Not rs.EOF
         If rsaux.State = 1 Then
            rsaux.Close
         End If
         var_planta_id = rs("VCHA_TRA_PLANTA_ORIGEN").Value
         If var_clave_movimiento = "DPL" Then
            rsaux.Open "select *  from tb_plantas where vcha_pla_planta_id = '" + rs("VCHA_TRA_PLANTA_DESTINO").Value + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
         Else
            If var_clave_movimiento = "EP" Then
                rsaux.Open "select VCHA_ALM_NOMBRE vcha_pla_descripc , VCHA_PLA_PLANTA_ID " & _
                            "from TB_TRANSITO_ALMACENES " & _
                            "where VCHA_ALM_ALMACEN_ID = '" + var_planta_id + "'", _
                        cnnoracle, _
                        adOpenDynamic, _
                        adLockOptimistic
            Else
                rsaux.Open "select *  from tb_plantas where vcha_pla_planta_id = '" + var_planta_id + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
            End If
        End If
         If Not rsaux.EOF Then
            If var_clave_movimiento = "EP" Or var_clave_movimiento = "EI" Then
               var_nota = ""
               var_i = 0
               For var_j = 1 To Len(rs!VCHA_TRA_NOTA_ENVIO)
                   If Mid(rs!VCHA_TRA_NOTA_ENVIO, var_j, 1) = "_" Then
                      var_i = var_j
                   End If
               Next var_j
               var_nota = Mid(rs!VCHA_TRA_NOTA_ENVIO, var_i + 1, 50)
            Else
               'var_nota = rs!VCHA_TRA_NOTA_ENVIO
                strNotaTraspaso = Mid(rs!VCHA_TRA_NOTA_ENVIO, InStr(rs!VCHA_TRA_NOTA_ENVIO, "_") + 1, 50)
                For var_j = 1 To Len(strNotaTraspaso)
                    If IsNumeric(Mid(strNotaTraspaso, var_j, 1)) = True Then
                        var_nota = var_nota & Mid(strNotaTraspaso, var_j, 1)
                    End If
                Next
            End If
            If var_clave_movimiento = "EI" Then
               var_nota = UCase(IIf(IsNull(rs!vcha_Ser_Serie_id), "", rs!vcha_Ser_Serie_id)) + var_nota
            End If
            Set list_item = lv_notas.ListItems.Add(, , Mid(var_nota, InStr(var_nota, "_") + 1, 50))
            'Set list_item = lv_notas.ListItems.Add(, , var_nota, InStr(var_nota, "_") + 1, 50))
            list_item.SubItems(1) = IIf(IsNull(rsaux!vcha_pla_descripc), "", rsaux!vcha_pla_descripc)
            list_item.SubItems(2) = Format(IIf(IsNull(rs!CANTIDAD_ENVIADA), 0, rs!CANTIDAD_ENVIADA), "###,###,##0.00")
            list_item.SubItems(3) = Format(IIf(IsNull(rs!CANTIDAD_RECIBIDA), 0, rs!CANTIDAD_RECIBIDA), "###,###,##0.00")
            list_item.SubItems(4) = IIf(IsNull(rs!VCHA_TRA_PLANTA_ORIGEN), "", rs!VCHA_TRA_PLANTA_ORIGEN)
            list_item.SubItems(5) = IIf(IsNull(rs!VCHA_TRA_NOTA_ENVIO), "", rs!VCHA_TRA_NOTA_ENVIO)
            list_item.SubItems(6) = CInt(Date - CDate(IIf(IsNull(rs!DATE_tRA_FECHA_ENVIO), Date, rs!DATE_tRA_FECHA_ENVIO)))
            list_item.SubItems(7) = IIf(IsNull(rs!VCHA_TRA_PLANTA_DESTINO), "", rs!VCHA_TRA_PLANTA_DESTINO)
            If var_clave_movimiento = "EP" Then
                list_item.SubItems(7) = IIf(IsNull(rs!PLANTA_ORIGEN), "", rs!PLANTA_ORIGEN)
            End If
         End If
         var_nota = ""
         rsaux.Close
         rs.MoveNext:
         numero_items_licencias = numero_items_licencias + 1
    Wend
    rs.Close

End Sub

Private Sub lv_notas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_notas, ColumnHeader)
End Sub

Private Sub lv_notas_KeyPress(KeyAscii As Integer)
   Dim var_zzz As Integer
   If KeyAscii = 13 Then
      var_zzz = 0
      If Me.lv_notas.ListItems.Count > 0 Then
         var_clave_unidad_origen = Me.lv_notas.selectedItem.SubItems(4)
         var_nota_traspasos = Me.lv_notas.selectedItem
         var_nota_traspasos_transito = Me.lv_notas.selectedItem.SubItems(5)
         
         
         Select Case var_clave_movimiento
            Case "EP"
                rs.Open "Select numb_tra_consecutivo, VCHA_TRA_NOTA_ENVIO , " & _
                                "VCHA_TRA_ALMACEN_ORIGEN, VCHA_ART_ARTICULO_ORIGEN, " & _
                                "NUMB_TRA_CANTIDAD_ENVIADA, VCHA_TRA_REFERENCIA1, NVL(VCHA_TRA_CONTENEDOR_ID,' ') VCHA_TRA_CONTENEDOR_ID " & _
                            "from tb_transito " & _
                            "where VCHA_TRA_NOTA_ENVIO = '" & var_nota_traspasos_transito & "' ", _
                        cnnoracle, _
                        adOpenDynamic, _
                        adLockOptimistic
                
                For fila = 1 To rs.RecordCount
                    rsaux.Open "Update tb_archivo_comparacion " & _
                                "set inte_com_consecutivo =" & rs("numb_tra_consecutivo").Value & ", " & _
                                    "vcha_com_referencia_almacen = '" & rs("VCHA_TRA_ALMACEN_ORIGEN").Value & "'" & _
                                "where vcha_art_articulo_id ='" & rs("VCHA_ART_ARTICULO_ORIGEN").Value & "' " & _
                                "and VCHA_COM_REFERENCIA ='" & var_nota_traspasos_transito & "' " & _
                                "and inte_com_lote ='" & rs("VCHA_TRA_REFERENCIA1").Value & "' and vcha_com_caja = '" & rs("VCHA_TRA_CONTENEDOR_ID").Value & "'", _
                            cnn, _
                            adOpenDynamic, _
                            adLockOptimistic
                    rs.MoveNext
                Next
                rs.Close
                Unload Me
                Exit Sub
            Case "DPL"
                rs.Open "select * " & _
                        "from tb_transito " & _
                        "where vcha_tra_nota_envio = '" + var_nota_traspasos_transito + "' " & _
                        "AND FLOA_TRA_CANTIDAD_ENVIADA > FLOA_TRA_CANTIDAD_RECIBIDA ", _
                    cnn_admcdindustrial, _
                    adOpenDynamic, _
                    adLockOptimistic
            Case Else
                rs.Open "select * " & _
                        "from tb_Archivo_comparacion " & _
                        "where vcha_com_referencia = '" + var_nota_traspasos_transito + "' " & _
                        "and vcha_mov_movimiento_id = '" + var_clave_movimiento + "' " & _
                        "and vcha_Emp_empresa_id = '" + var_empresa + "'", _
                    cnn, _
                    adOpenDynamic, _
                    adLockOptimistic
        End Select
         
         If var_clave_movimiento = "DPL" Then
            
         Else
         
            
         End If
            
         If rs.EOF Then
            
            If var_clave_movimiento = "ETA" Or var_clave_movimiento = "DPL" Then
               If var_clave_usuario_global = "U0000000212" Or var_clave_usuario_global = "U0000000211" Then
                  var_almacen_nota = "ABPT"
               Else
                  If var_unidad_organizacional = "12" Then
                     rsaux.Open "select * from tb_Almacenes where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND vcha_alm_almacen_id = '8'", cnn, adOpenDynamic, adLockOptimistic
                     var_almacen_nota = IIf(IsNull(rsaux!VCHA_ALM_ALMACEN_ID), "", rsaux!VCHA_ALM_ALMACEN_ID)
                  Else
                     If var_unidad_organizacional = "28" Then
                        var_almacen_nota = "MPCOC"
                     Else
                        'MsgBox "select * from tb_Almacenes where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND INTE_INT_ENTRADA_TRASPASO_PLANTAS =  1"
                        rsaux.Open "select * from tb_Almacenes where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' AND INTE_INT_ENTRADA_TRASPASO_PLANTAS =  1", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux.EOF Then
                           If var_clave_unidad_planta = "RETEX" Then
                              var_almacen_nota = "RETEX"
                           Else
                              var_almacen_nota = IIf(IsNull(rsaux!VCHA_ALM_ALMACEN_ID), "", rsaux!VCHA_ALM_ALMACEN_ID)
                           End If
                        Else
                           var_almacen_nota = ""
                        End If
                        
                     End If
                  End If
               End If
            Else
               If var_clave_usuario_global = "U0000000212" Or var_clave_usuario_global = "U0000000211" Then
                  var_almacen_nota = "PTMU"
               Else
                  rsaux.Open "select * from tb_Almacenes where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn, adOpenDynamic, adLockOptimistic
                  If Not rsaux.EOF Then
                     var_almacen_nota = IIf(IsNull(rsaux!VCHA_ALM_ALMACEN_ID), "", rsaux!VCHA_ALM_ALMACEN_ID)
                  Else
                     var_almacen_nota = ""
                  End If
               End If
            End If
            If rsaux.State = 1 Then
               rsaux.Close
            End If
            If var_almacen_nota <> "" Then
               If rsaux.State = 1 Then
                  rsaux.Close
               End If
               If var_clave_movimiento = "DPL" Then
               
                    rsaux.Open "select * from tb_plantas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                Else
                    rsaux.Open "select * from tb_plantas where vcha_uor_unidad_id = '" + var_unidad_organizacional + "' and vcha_pla_planta_id = '" + var_planta_transito_global + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                End If
               If Not rsaux.EOF Then
               
                  var_planta_destino_archivo = IIf(IsNull(rsaux!vcha_pla_planta_id), "", rsaux!vcha_pla_planta_id)
               End If
               rsaux.Close
               rsaux.Open "select * from tb_Transito where vcha_tra_nota_envio = '" + var_nota_traspasos_transito + "' and vcha_tra_planta_destino = '" + var_planta_destino_archivo + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
               var_posible_articulos = True
               While Not rsaux.EOF
                     If rsaux1.State = 1 Then
                        rsaux1.Close
                     End If
                     If var_clave_unidad_planta = "PTTEX" Then
                        rsaux1.Open "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "' and len(vcha_art_articulo_id) = 12 and vcha_Art_Articulo_id <> '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                        If Not rsaux1.EOF Then
                           rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If rsaux2.EOF Then
                              var_posible_articulos = False
                           End If
                           rsaux2.Close
                        Else
                           rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If rsaux2.EOF Then
                              var_posible_articulos = False
                           End If
                           rsaux2.Close
                        End If
                        rsaux1.Close
                     Else
                        If var_clave_unidad_origen = "PTTEX" Or var_clave_unidad_origen = "TEXMP" Or var_clave_unidad_origen = "RETEX" Or var_clave_unidad_planta = "9" Then
                           rsaux1.Open "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                           If Not rsaux1.EOF Then
                              rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              If rsaux2.EOF Then
                                 var_posible_articulos = False
                              End If
                              rsaux2.Close
                           Else
                              var_posible_articulos = False
                           End If
                           rsaux1.Close
                        Else
                           If (var_clave_movimiento = "EP" Or var_clave_movimiento = "ETA") And var_empresa = "02" Then
                              If rsaux9.State = 1 Then
                                 rsaux9.Close
                              End If
                              rsaux9.Open "select * from tb_Articulos where vcha_Art_Articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                              If Not rsaux9.EOF Then
                                 var_posible_articulos = True
                              Else
                                 rsaux1.Open "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If Not rsaux1.EOF Then
                                    rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If rsaux2.EOF Then
                                       var_posible_articulos = False
                                    End If
                                    rsaux2.Close
                                 Else
                                    rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If rsaux2.EOF Then
                                       var_posible_articulos = False
                                    End If
                                    rsaux2.Close
                                 End If
                                 rsaux1.Close
                              End If
                           Else
                              If var_clave_unidad_origen = "0" And var_empresa = "18" Then
                                    rsaux2.Open "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If rsaux2.EOF Then
                                       var_posible_articulos = False
                                    End If
                                    rsaux2.Close
                              Else
                                 rsaux1.Open "SELECT * FROM TB_ARTICULOS WHERE VCHA_aRT_aRTICULO_ID = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                 If rsaux1.EOF Then
                                    'MsgBox "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "'"
                                    rsaux2.Open "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If rsaux2.EOF Then
                                       var_posible_articulos = False
                                    End If
                                    rsaux2.Close
                                 End If
                                 rsaux1.Close
                              End If
                           End If
                        End If
                     End If
                     rsaux.MoveNext
               Wend
               If rsaux.RecordCount > 0 Then
                  rsaux.MoveFirst
                  If var_posible_articulos = True Then
                     var_consecutivo = 0
                     var_nota_tr = IIf(IsNull(rsaux!VCHA_TRA_NOTA_ENVIO), "0", rsaux!VCHA_TRA_NOTA_ENVIO)
                     VAR_NUMERO_NOTA_TR = ""
                     VAR_SI_NOTA = 0
                     For var_j = 1 To Len(var_nota_tr)
                         If Mid(var_nota_tr, var_j, 1) = "_" Then
                            VAR_SI_NOTA = 1
                         End If
                         If VAR_SI_NOTA = 1 Then
                            If Mid(var_nota_tr, var_j, 1) <> "_" And IsNumeric(Mid(var_nota_tr, var_j, 1)) Then
                               VAR_NUMERO_NOTA_TR = VAR_NUMERO_NOTA_TR + Mid(var_nota_tr, var_j, 1)
                            End If
                         End If
                     Next var_j
                     If var_clave_movimiento = "EP" Then
                        VAR_NUMERO_NOTA_TR = Mid(VAR_NUMERO_NOTA_TR, 6, 10)
                     End If
                     
                     If VAR_NUMERO_NOTA_TR <> "" Then
                        While Not rsaux.EOF
                              var_consecutivo = var_consecutivo + 1
                              If rsaux11.State = 1 Then
                                 rsaux11.Close
                              End If
                              rsaux11.Open "SELECT * FROM TB_PLANTAS WHERE VCHA_PLA_PLANTA_ID = '" + Me.lv_notas.selectedItem.SubItems(4) + "'", cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                              If Not rsaux11.EOF Then
                                 VAR_PLANTA_ORIGEN_PROVEEDOR = IIf(IsNull(rsaux11!vcha_uor_unidad_id), "", rsaux11!vcha_uor_unidad_id)
                              Else
                                 VAR_PLANTA_ORIGEN_PROVEEDOR = ""
                              End If
                              rsaux11.Close
                              
                              If var_clave_unidad_origen = "PTTEX" Or var_clave_unidad_origen = "TEXMP" Then
                                 If var_clave_movimiento = "EP" Then
                                    rsaux3.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux3.EOF Then
                                       var_codigo = rsaux3!vcha_Art_Articulo_id
                                    End If
                                    rsaux3.Close
                                 Else
                                    rsaux1.Open "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux1.EOF Then
                                       rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux2.EOF Then
                                          var_codigo = rsaux2!vcha_Art_Articulo_id
                                       End If
                                       rsaux2.Close
                                    End If
                                    rsaux1.Close
                                 End If
                              Else
                                 If (var_clave_movimiento = "EP" And var_empresa = "02") Or var_clave_unidad_planta = "9" Then
                                    rsaux1.Open "select * from tb_Equivalencias where vcha_equ_codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    If Not rsaux1.EOF Then
                                       rsaux2.Open "select * from tb_Articulos where vcha_Art_articulo_id = '" + rsaux1!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                       If Not rsaux2.EOF Then
                                          var_codigo = rsaux2!vcha_Art_Articulo_id
                                       End If
                                       rsaux2.Close
                                    End If
                                    rsaux1.Close
                                 Else
                                    If var_empresa = "18" Then
                                       rsaux1.Open "select * from tb_Articulos where vcha_Art_Articulo_id = '" + rsaux!vcha_Art_Articulo_id + "' and len(vcha_Art_articulo_id) = 12", cnn, adOpenDynamic, adLockOptimistic
                                    Else
                                       rsaux1.Open "select * from tb_Articulos where vcha_Art_Articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                    End If
                                    If Not rsaux1.EOF Then
                                       var_codigo = rsaux!vcha_Art_Articulo_id
                                    Else
                                       If var_empresa = "18" Then
                                          rsaux2.Open "select * from tb_equivalencias where vcha_equ_Codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "' and len(vcha_art_articulo_id) = 12", cnn, adOpenDynamic, adLockOptimistic
                                       Else
                                          rsaux2.Open "select * from tb_equivalencias where vcha_equ_Codigo_equivalente = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
                                       End If
                                       If Not rsaux2.EOF Then
                                          var_codigo = rsaux2!vcha_Art_Articulo_id
                                       End If
                                       rsaux2.Close
                                    End If
                                    rsaux1.Close
                                    
                                 End If
                              End If
                              If var_empresa = "02" Then
                              Else
                                 If VAR_PLANTA_ORIGEN_PROVEEDOR = "82" Or VAR_PLANTA_ORIGEN_PROVEEDOR = "80" Then
                                    var_almacen_nota = "MPCOC"
                                 End If
                                  If VAR_PLANTA_ORIGEN_PROVEEDOR = "79" Or VAR_PLANTA_ORIGEN_PROVEEDOR = "83" Then
                                     var_almacen_nota = "MPEDR"
                                  End If
                                  If VAR_PLANTA_ORIGEN_PROVEEDOR = "81" Or VAR_PLANTA_ORIGEN_PROVEEDOR = "84" Then
                                    var_almacen_nota = "MPCOL"
                                 End If
                              End If
                                                      
                              var_cadena = "insert into tb_Archivo_comparacion (vcha_emp_empresa_id, vcha_uor_unidad_id, vcha_Alm_almacen_id, vcha_mov_movimiento_id, inte_com_numero, dtim_com_Fecha, char_com_tipo_proveedor, vcha_com_proveedor, vcha_Art_articulo_id, floa_com_costo, floa_com_Cantidad_enviada, floa_com_cantidad_recibida, vcha_com_transporto, vcha_com_referencia, inte_com_lote, inte_com_consecutivo, VCHA_COM_REFERENCIA_TRANSITO)"
                              var_cadena = var_cadena + " values ('" + var_empresa + "','" + var_unidad_organizacional + "', '" + var_almacen_nota + "','" + var_clave_movimiento + "'," + VAR_NUMERO_NOTA_TR + ",GETDATE(),'U','" + VAR_PLANTA_ORIGEN_PROVEEDOR + "','" + var_codigo + "'," + CStr(rsaux!floa_Tra_Costo) + "," + CStr(rsaux!FLOA_TRA_CANTIDAD_ENVIADA) + ",0,'','" + Me.lv_notas.selectedItem.SubItems(5) + "',0," + CStr(var_consecutivo) + ",'" + Me.lv_notas.selectedItem.SubItems(5) + "')"
                              rsaux1.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
                              var_cadena = "update tb_transito set vcha_Art_articulo_recivo = '" + var_codigo + "' where inte_Tra_consecutivo = " + CStr(rsaux!inte_tra_consecutivo)
                              rsaux1.Open var_cadena, cnn_admcdindustrial, adOpenDynamic, adLockOptimistic
                              rsaux.MoveNext
                        Wend
                     Else
                        MsgBox "Número de nota incorrecto", vbOKOnly, "ATENCION"
                     End If
                  Else
                     
                     MsgBox "La nota contiene artículos que no estan dados de alta en el almacén", vbOKOnly, "ATENCION"
                     frmarticulos_transito.Show 1
                     var_zzz = 1
                  End If
               Else
                  MsgBox "El archivo no contiene información", vbOKOnly, "ATENCION"
               End If
               rsaux.Close
            Else
               MsgBox "No existe un almacen destino", vbOKOnly, "ATENCION"
            End If
         Else
            If var_clave_movimiento = "DPL" Then
                var_str_nota_envio = var_nota_traspasos_transito
            Else
                var_str_nota_envio = ""
            End If
         End If
         rs.Close
         If var_zzz = 0 Then
            Unload Me
         End If
      End If
   End If
   If KeyAscii = 27 Then
      var_nota_traspasos = ""
      Unload Me
   End If
End Sub
