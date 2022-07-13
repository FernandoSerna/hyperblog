VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_creacion_palets 
   Appearance      =   0  'Flat
   Caption         =   "Creación de palets"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      Picture         =   "frmoracle_creacion_palets.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cerrar Caja e Imprimir las Etiquetas"
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   695
      Picture         =   "frmoracle_creacion_palets.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cerrar Caja e Imprimir las Etiquetas"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_buscar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      Picture         =   "frmoracle_creacion_palets.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   50
      TabIndex        =   6
      Top             =   495
      Width           =   5655
      Begin VB.TextBox txt_pedido 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   12
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox txt_bulto 
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   8
         Top             =   945
         Width           =   2715
      End
      Begin VB.TextBox txt_palet 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   7
         Top             =   255
         Width           =   2715
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Pedido:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   660
         Width           =   540
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Bulto:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label lab_paises 
         AutoSize        =   -1  'True
         Caption         =   "Palet:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   315
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   50
      TabIndex        =   4
      Top             =   1830
      Width           =   5655
      Begin MSComctlLib.ImageList icono_encabezado 
         Left            =   0
         Top             =   0
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
               Picture         =   "frmoracle_creacion_palets.frx":0306
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmoracle_creacion_palets.frx":0BE0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lv_bultos 
         Height          =   4725
         Left            =   45
         TabIndex        =   5
         Top             =   150
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   8334
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bulto"
            Object.Width           =   9701
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   45
      Picture         =   "frmoracle_creacion_palets.frx":14BA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5325
      Picture         =   "frmoracle_creacion_palets.frx":15BC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   50
      TabIndex        =   11
      Top             =   315
      Width           =   5655
   End
End
Attribute VB_Name = "frmoracle_creacion_palets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_consecutivo As Integer

Private Sub cmd_buscar_Click()
   Me.txt_bulto = ""
   Me.txt_palet = ""
   Me.txt_pedido = ""
   Me.txt_bulto.Enabled = False
   Me.lv_bultos.ListItems.Clear
   Me.txt_palet.Enabled = True
   Me.txt_palet.SetFocus
End Sub

Private Sub cmd_imprimir_Click()
   If Me.lv_bultos.ListItems.Count > 0 Then
      var_si = MsgBox("¿Desea generar la etiqueta?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la impresión de la etiqueta", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If cnn_icg_usa.State = 1 Then
               cnn_icg_usa.Close
            End If
            cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=ICGUsa2014;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=sid;Data Source=sqlcedishou.VIANNEYcatalog.COM"
            'cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=sidcedishou;Data Source=sqlquezada2"
            rs.Open "select * from tb_oracle_pedidos_asignados_embarques where pedido = '" + Me.txt_pedido + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If rs!Agente = "1016" Then
                  var_cliente = "CLIENTE: " + IIf(IsNull(rs!NOMBRE_AGENTE), "", rs!NOMBRE_AGENTE)
               Else
                  var_cliente = "CLIENTE: " + IIf(IsNull(rs!Cliente), "", rs!Cliente)
               End If
            Else
               var_cliente = ""
            End If
            rs.Close
            rs.Open "delete tb_oracle_palets where palet = '" + Me.txt_palet + "'", cnn_icg_usa, adOpenDynamic, adLockOptimistic
            For var_j = 1 To Me.lv_bultos.ListItems.Count
                Me.lv_bultos.ListItems.Item(var_j).Selected = True
                rs.Open "insert into tb_oracle_palets (palet, caja, estatus, pedido) values ('" + Me.txt_palet + "','" + Me.lv_bultos.selectedItem + "','','" + Me.txt_pedido + "')", cnn_icg_usa, adOpenDynamic, adLockOptimistic
            Next var_j
            
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set A = fs.CreateTextFile(App.Path + "\palet.txt", True)
            A.writeline ("")
            A.writeline ("US")
            A.writeline ("N")
            A.writeline ("q816")
            A.writeline ("Q1015,20+0")
            A.writeline ("S2")
            A.writeline ("D8")
            A.writeline ("ZT")
            A.writeline ("TTh:m")
            A.writeline ("TDy2.mn.dd")
            var_numero_etiqueta = 1
            For var_j = 1 To Me.lv_bultos.ListItems.Count
                If var_numero_etiqueta > 24 Then
                   var_numero_etiqueta = 1
                   
                   A.writeline ("A320,20,1,4,2,2,N,""PEDIDO: " + Me.txt_pedido + """")
                   A.writeline ("A220,20,1,4,2,1,N,""" + Mid(var_cliente, 1, 60) + """")
                   A.writeline ("B160,20,1,3,4,8,101,B,""" + Me.txt_palet + """")
                   A.writeline ("P1")
                   A.writeline ("")
                   A.writeline ("US")
                   A.writeline ("N")
                   A.writeline ("q816")
                   A.writeline ("Q1015,20+0")
                   A.writeline ("S2")
                   A.writeline ("D8")
                   A.writeline ("ZT")
                   A.writeline ("TTh:m")
                   A.writeline ("TDy2.mn.dd")
                
                End If
                Me.lv_bultos.ListItems.Item(var_j).Selected = True
                If var_numero_etiqueta = 1 Then
                   A.writeline ("A782,20,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 2 Then
                   A.writeline ("A696,20,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 3 Then
                   A.writeline ("A627,20,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 4 Then
                   A.writeline ("A554,20,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 5 Then
                   A.writeline ("A475,20,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 6 Then
                   A.writeline ("A390,20,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 7 Then
                   A.writeline ("A782,250,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 8 Then
                   A.writeline ("A696,250,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 9 Then
                   A.writeline ("A627,250,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 10 Then
                   A.writeline ("A554,250,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 11 Then
                   A.writeline ("A475,250,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                
                
                If var_numero_etiqueta = 14 Then
                   A.writeline ("A390,250,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 15 Then
                   A.writeline ("A782,480,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 16 Then
                   A.writeline ("A696,480,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 17 Then
                   A.writeline ("A627,480,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 18 Then
                   A.writeline ("A554,480,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 19 Then
                   A.writeline ("A475,480,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 20 Then
                   A.writeline ("A390,480,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                
                If var_numero_etiqueta = 21 Then
                   A.writeline ("A782,710,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 22 Then
                   A.writeline ("A696,710,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 23 Then
                   A.writeline ("A627,710,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 24 Then
                   A.writeline ("A554,710,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 25 Then
                   A.writeline ("A475,710,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                If var_numero_etiqueta = 26 Then
                   A.writeline ("A390,710,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                
                
                If var_numero_etiqueta = 12 Then
                   A.writeline ("A390,710,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                
                If var_numero_etiqueta = 13 Then
                   A.writeline ("A475,710,1,4,2,1,N,""" + Me.lv_bultos.selectedItem + """")
                End If
                
                
                var_numero_etiqueta = var_numero_etiqueta + 1
            Next var_j
            If var_numero_etiqueta < 26 Then
               A.writeline ("A320,20,1,4,2,2,N,""PEDIDO: " + Me.txt_pedido + """")
               A.writeline ("A220,20,1,4,2,1,N,""" + Mid(var_cliente, 1, 60) + """")
               A.writeline ("B160,20,1,3,4,8,101,B,""" + Me.txt_palet + """")
               A.writeline ("P1")
            End If
            Open (App.Path & "\net_use.bat") For Output As #3
            var_archivo = App.Path & "\net_use.bat"
            Print #3, "net use lpt1 /delete"
            Print #3, "net use lpt1 \\" + fun_NombrePc + "\zebra2 /persistent:yes"
            Close #3
            x = Shell(var_archivo, vbHide)
            
            
            Open (App.Path & "\palet.bat") For Output As #2
            var_archivo = App.Path & "\palet.bat"
            Print #2, "copy " + App.Path + "\palet.txt lpt1"
            Close #2
            x = Shell(var_archivo, vbHide)
            rs.Open "update tb_oracle_palets set estatus = 'I' where palet = '" + Me.txt_palet + "'", cnn, adOpenDynamic, adLockOptimistic
            Me.txt_bulto.Enabled = False
         End If
      End If
   Else
      MsgBox "No existen bultos en el palet", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.lv_bultos.ListItems.Clear
   Me.txt_palet = ""
   Me.txt_bulto = ""
   Me.txt_pedido = ""
   Me.txt_bulto.Enabled = True
   Me.txt_bulto.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   If Me.lv_bultos.ListItems.Count > 0 Then
      var_si = MsgBox("¿Desea generar la etiqueta?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la impresión de la etiqueta", vbYesNo, "ATENCION")
         If var_si = 6 Then
            If cnn_icg_usa.State = 1 Then
               cnn_icg_usa.Close
            End If
            cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=ICGUsa2014;Persist Security Info=True;User ID=ICGAdmin;Initial Catalog=sid;Data Source=sqlcedishou.VIANNEYcatalog.COM"
            'cnn_icg_usa.Open "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=sidcedishou;Data Source=sqlquezada2"
            If rs.State = 1 Then
               rs.Close
            End If
            rs.Open "select * from tb_oracle_palets", cnn, adOpenDynamic, adLockOptimistic
            While Not rs.EOF
                  rsaux1.Open "insert into tb_oracle_palets (palet, caja, estatus, pedido) values ('" + rs!PALET + "','" + rs!Caja + "','I'," + CStr(rs!pedido) + ")", cnn_icg_usa, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
         End If
      End If
   Else
      MsgBox "No existen bultos en el palet", vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub txt_bulto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txt_palet = "" Then
         var_embarque = Mid(Me.txt_bulto, 2, 6)
         rsaux1.Open "select * from tb_oracle_cajas_aduana  where caja = '" + Me.txt_bulto + "' AND EMBARQUE = " + var_embarque, cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_pedido = rsaux1!pedido
            'cnn.Open "Provider=SQLOLEDB.1;Password=" & parametros(3) & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & parametros(1) & ";Data Source=" & parametros(0)
            rs.Open "SELECT * FROM TB_ORACLE_CONSECUTIVO_PALETS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!PALET), 0, rs!PALET) + 1
               rsaux.Open "UPDATE TB_ORACLE_CONSECUTIVO_PALETS set PALET = PALET + 1", cnn, adOpenDynamic, adLockOptimistic
            Else
               var_consecutivo = 1
               rsaux.Open "INSERT INTO TB_ORACLE_CONSECUTIVO_PALETS (PALET) VALUES (1)", cnn, adOpenDynamic, adLockOptimistic
            End If
            rs.Close
            If Len(CStr(var_consecutivo)) = 1 Then
               VAR_PALET = "PL000000000" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 2 Then
               VAR_PALET = "PL00000000" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 3 Then
               VAR_PALET = "PL0000000" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 4 Then
               VAR_PALET = "PL000000" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 5 Then
               VAR_PALET = "PL00000" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 6 Then
               VAR_PALET = "PL0000" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 7 Then
               VAR_PALET = "PL000" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 8 Then
               VAR_PALET = "PL00" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 9 Then
               VAR_PALET = "PL0" + CStr(var_consecutivo)
            End If
            If Len(CStr(var_consecutivo)) = 10 Then
               VAR_PALET = "PL" + CStr(var_consecutivo)
            End If
            rsaux3.Open "select * from tb_oracle_palets where caja = '" + Me.txt_bulto + "'", cnn, adOpenDynamic, adLockOptimistic
            If rsaux3.EOF Then
               rsaux2.Open "insert into tb_oracle_palets (palet, caja, pedido, estatus) values ('" + VAR_PALET + "','" + Me.txt_bulto + "','" + Me.txt_pedido + "','')", cnn, adOpenDynamic, adLockOptimistic
               Set list_item = lv_bultos.ListItems.Add(, , Trim(Me.txt_bulto))
               Me.txt_palet = VAR_PALET
            Else
               frmmensaje.lbl_articulo = ""
               frmmensaje.lbl_mensaje = "El bulto ya fue leido en el palet " + rsaux3!PALET
               frmmensaje.Show 1
            End If
            rsaux3.Close
         Else
            frmmensaje.lbl_articulo = ""
            frmmensaje.lbl_mensaje = "El bulto no existe"
            frmmensaje.Show 1
         End If
         rsaux1.Close
         Me.txt_bulto = ""
      Else
         rsaux1.Open "select * from tb_oracle_cajas_aduana  where caja = '" + Me.txt_bulto + "' AND EMBARQUE = '" + Mid(Me.txt_bulto, 2, 6) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            If rsaux1!pedido = Me.txt_pedido Then
               rsaux3.Open "select * from tb_oracle_palets where caja = '" + Me.txt_bulto + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux3.EOF Then
                  rsaux2.Open "insert into tb_oracle_palets (palet, caja, pedido, estatus) values ('" + Me.txt_palet + "','" + Me.txt_bulto + "','" + Me.txt_pedido + "','')", cnn, adOpenDynamic, adLockOptimistic
                  Set list_item = lv_bultos.ListItems.Add(, , Trim(Me.txt_bulto))
               Else
                  frmmensaje.lbl_articulo = ""
                  frmmensaje.lbl_mensaje = "El bulto ya fue leido en el palet " + rsaux3!PALET
                  frmmensaje.Show 1
               End If
               rsaux3.Close
            Else
               frmmensaje.lbl_articulo = ""
               frmmensaje.lbl_mensaje = "El bulto no pertenece al pedido " + CStr(rsaux1!pedido)
               frmmensaje.Show 1
            End If
         Else
            frmmensaje.lbl_articulo = ""
            frmmensaje.lbl_mensaje = "El bulto no existe"
            frmmensaje.Show 1
         End If
         rsaux1.Close
         Me.txt_bulto = ""
      End If
   End If
End Sub

Private Sub txt_palet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If IsNumeric(Me.txt_palet) Then
          var_consecutivo = CDbl(Me.txt_palet)
          If Len(CStr(var_consecutivo)) = 1 Then
             VAR_PALET = "PL000000000" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 2 Then
             VAR_PALET = "PL00000000" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 3 Then
             VAR_PALET = "PL0000000" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 4 Then
             VAR_PALET = "PL000000" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 5 Then
             VAR_PALET = "PL00000" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 6 Then
             VAR_PALET = "PL0000" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 7 Then
             VAR_PALET = "PL000" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 8 Then
             VAR_PALET = "PL00" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 9 Then
             VAR_PALET = "PL0" + CStr(var_consecutivo)
          End If
          If Len(CStr(var_consecutivo)) = 10 Then
             VAR_PALET = "PL" + CStr(var_consecutivo)
          End If
          rsaux3.Open "select * from tb_oracle_palets where PALET = '" + VAR_PALET + "'", cnn, adOpenDynamic, adLockOptimistic
          If Not rsaux3.EOF Then
             Me.txt_palet = rsaux3!PALET
             If rsaux3!estatus = "" Then
                Me.txt_bulto.Enabled = True
             Else
                Me.txt_bulto.Enabled = False
             End If
             Me.txt_pedido = rsaux3!pedido
             While Not rsaux3.EOF
                   Set list_item = lv_bultos.ListItems.Add(, , Trim(rsaux3!Caja))
                   rsaux3.MoveNext
             Wend
             If Me.txt_bulto.Enabled = True Then
                Me.txt_bulto.SetFocus
                Me.txt_palet.Enabled = False
             End If
          Else
             MsgBox "El palet no existe", vbOKOnly, "ATENCION"
             Me.txt_palet = ""
             Me.txt_palet.Enabled = False
          End If
          rsaux3.Close
          Me.txt_palet.Enabled = False
       Else
          rsaux3.Open "select * from tb_oracle_palets where PALET = '" + Me.txt_palet + "'", cnn, adOpenDynamic, adLockOptimistic
          If Not rsaux3.EOF Then
             Me.txt_palet = rsaux3!PALET
             If rsaux3!estatus = "" Then
                Me.txt_bulto.Enabled = True
             Else
                Me.txt_bulto.Enabled = False
             End If
             Me.txt_pedido = rsaux3!pedido
             While Not rsaux3.EOF
                   Set list_item = lv_bultos.ListItems.Add(, , Trim(rsaux3!Caja))
                   rsaux3.MoveNext
             Wend
             If Me.txt_bulto.Enabled = True Then
                Me.txt_bulto.SetFocus
             End If
          Else
             MsgBox "El palet no existe", vbOKOnly, "ATENCION"
             Me.txt_palet = ""
          End If
          rsaux3.Close
          Me.txt_palet.Enabled = False
       End If
    End If
End Sub
