VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmoracle_validador_cubicaje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validador de cubicaje"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   17070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   75
      TabIndex        =   13
      Top             =   9405
      Width           =   16905
      Begin VB.TextBox txt_porcentaje_total 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   15570
         TabIndex        =   32
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txt_total_pedido 
         Height          =   285
         Left            =   3000
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_volumen_total 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13080
         TabIndex        =   28
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txt_porcentaje 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10410
         TabIndex        =   21
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txt_volumen_pedidos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7050
         TabIndex        =   19
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox txt_volumen_unidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4290
         TabIndex        =   17
         Top             =   180
         Width           =   990
      End
      Begin VB.ComboBox cmb_tipo_unidades 
         Height          =   315
         ItemData        =   "frmoracle_validador_cubicaje.frx":0000
         Left            =   960
         List            =   "frmoracle_validador_cubicaje.frx":0002
         TabIndex        =   15
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "% Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   14640
         TabIndex        =   33
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "% de Ocupacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8640
         TabIndex        =   29
         Top             =   240
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vol. Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11925
         TabIndex        =   20
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vol. Pedidos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5445
         TabIndex        =   18
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vol. Unidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame frmbusqueda_pedido 
      Height          =   3900
      Left            =   630
      TabIndex        =   6
      Top             =   360
      Width           =   6525
      Begin VB.TextBox txt_ruta 
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3390
         Width           =   6315
      End
      Begin VB.CommandButton cmd_buscar_pedido 
         Caption         =   "Cargar archivo"
         Height          =   465
         Left            =   3330
         TabIndex        =   10
         Top             =   2790
         Width           =   3060
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   90
         TabIndex        =   9
         Top             =   930
         Width           =   3150
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   3330
         Pattern         =   "*.xls"
         TabIndex        =   8
         Top             =   510
         Width           =   3075
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   105
         TabIndex        =   7
         Top             =   510
         Width           =   3135
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Caption         =   " Busqueda de pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   6465
      End
   End
   Begin VB.CommandButton cmd_generar_archivos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmoracle_validador_cubicaje.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Generar archivos"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_cargar_archivo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      Picture         =   "frmoracle_validador_cubicaje.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cargar archivo"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   16680
      Picture         =   "frmoracle_validador_cubicaje.frx":0418
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   75
      TabIndex        =   2
      Top             =   375
      Width           =   16950
   End
   Begin VB.Frame Frame1 
      Height          =   9045
      Left            =   75
      TabIndex        =   0
      Top             =   360
      Width           =   16920
      Begin VB.TextBox txt_total_surtir 
         Height          =   495
         Left            =   8040
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frm_cantidad_agregar 
         Height          =   1035
         Left            =   10605
         TabIndex        =   25
         Top             =   3705
         Width           =   1905
         Begin VB.TextBox txt_cantidad_agregar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   60
            TabIndex        =   26
            Top             =   510
            Width           =   1800
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000002&
            Caption         =   " Cantidad a agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   345
            Left            =   30
            TabIndex        =   27
            Top             =   120
            Width           =   1830
         End
      End
      Begin VB.Frame frm_cantidad_eliminar 
         Height          =   1035
         Left            =   11175
         TabIndex        =   22
         Top             =   3705
         Width           =   1905
         Begin VB.TextBox txt_cantidad_eliminar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   60
            TabIndex        =   23
            Top             =   510
            Width           =   1800
         End
         Begin VB.Label Labe 
            BackColor       =   &H000000FF&
            Caption         =   " Cantidad a eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   345
            Left            =   30
            TabIndex        =   24
            Top             =   120
            Width           =   1830
         End
      End
      Begin MSComctlLib.ListView lv_lista 
         Height          =   8805
         Left            =   45
         TabIndex        =   1
         Top             =   150
         Width           =   16800
         _ExtentX        =   29633
         _ExtentY        =   15531
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Archivo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   9172
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cantidad"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Disponible"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Surtir"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Voll. Disponible"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Volumen unitario"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Archivo Número"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Vol. Total"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmoracle_validador_cubicaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter
Dim var_renglon As Integer

Private Sub cmb_tipo_unidades_Change()
   
   Me.txt_porcentaje = 0
   Me.txt_volumen_unidad = 0
   rs.Open "select * from tb_oracle_tipo_unidades where tipo_unidad = '" + Me.cmb_tipo_unidades.Text + "'", cnn, adOpenDynamic, adLockBatchOptimistic
   If Not rs.EOF Then
      Me.txt_volumen_unidad = IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN)
   Else
      MsgBox "No existe la unidad seleccionada"
      Me.txt_volumen_unidad = 0
   End If
   rs.Close
   If IsNumeric(Me.txt_volumen_pedidos) Then
      var_porcentaje = (CDbl(Me.txt_volumen_pedidos) / CDbl(Me.txt_volumen_unidad)) / 100
      Me.txt_porcentaje = var_porcentaje
   End If
   
End Sub

Private Sub cmb_tipo_unidades_Click()
   Me.txt_porcentaje = 0
   Me.txt_porcentaje_total = 0
   Me.txt_volumen_unidad = 0
   'rs.Open "select * from tb_oracle_tipo_unidades where tipo_unidad = '" + Me.cmb_tipo_unidades.Text + "'", cnn, adOpenDynamic, adLockBatchOptimistic
   rs.Open "SELECT NOMBRE, VOLUMEN  FROM tb_oracle_transportes where nombre = '" + Me.cmb_tipo_unidades.Text + "'", cnn, adOpenDynamic, adLockBatchOptimistic
   If Not rs.EOF Then
      Me.txt_volumen_unidad = IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN)
   Else
      MsgBox "No existe la unidad seleccionada"
      Me.txt_volumen_unidad = 0
   End If
   rs.Close
   
   
   
   If IsNumeric(Me.txt_volumen_pedidos) Then
      var_porcentaje = (CDbl(Me.txt_volumen_pedidos) / CDbl(Me.txt_volumen_unidad)) * 100
      Me.txt_porcentaje = var_porcentaje
   End If
   If IsNumeric(Me.txt_volumen_total) Then
      var_porcentaje_total = (CDbl(Me.txt_volumen_total) / CDbl(Me.txt_volumen_unidad)) * 100
      Me.txt_porcentaje_total = var_porcentaje_total
   End If
End Sub

Private Sub cmb_tipo_unidades_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_volumen_unidad.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub cmd_buscar_pedido_Click()
   var_numero_archivo = 0
   If Me.lv_lista.ListItems.Count > 0 Then
      Me.lv_lista.ListItems.Item(Me.lv_lista.ListItems.Count).Selected = True
      var_numero_archivo = CDbl(Me.lv_lista.selectedItem.SubItems(8))
   End If
   var_numero_archivo = var_numero_archivo + 1
On Error GoTo SALIR:
   var_archivo = Replace(Me.File1.FileName, ".xls", "")
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & Me.txt_ruta
   rsaux2.Open "SELECT codigo, cantidad FROM [PLANTILLA PEDIDO$] ", strConnectionString
   If Not rsaux2.EOF Then
      
      While Not rsaux2.EOF
            Set list_item = Me.lv_lista.ListItems.Add(, , var_archivo)
            var_codigo = CStr(IIf(IsNull(rsaux2!codigo), "", rsaux2!codigo))
            If Len(var_codigo) = 4 Then
               var_codigo = "0000" + CStr(var_codigo)
            End If
            If Len(var_codigo) = 5 Then
               var_codigo = "000" + CStr(var_codigo)
            End If
            If Len(var_codigo) = 6 Then
               var_codigo = "00" + CStr(var_codigo)
            End If
            If Len(var_codigo) = 7 Then
               var_codigo = "0" + CStr(var_codigo)
            End If
            
            list_item.SubItems(1) = var_codigo
            strconsulta = "select description, to_number(nvl(a.unit_volume,'0')) as volumen from xxvia_system_items_b a where a.segment1 = ? and organization_id = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_codigo)
                 .Parameters.Append parametro
                 Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, var_unidad_organizacional)
                 .Parameters.Append parametro
            End With
            Set rsaux6 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            If Not rsaux6.EOF Then
               var_volumen = IIf(IsNull(rsaux6!VOLUMEN), 0, rsaux6!VOLUMEN)
               var_descripcion = IIf(IsNull(rsaux6!Description), "", rsaux6!Description)
            Else
               var_volumen = 0
               var_descripcion = ""
            End If
            rsaux6.Close

            strconsulta = "select * from Xxvia_vw_existencias_inv where organization_id = 93 and subinventory_code = 'CDI_ALMPT' and segment1 = ?"
            With comandoORA
                 .ActiveConnection = cnnoracle_4
                 .CommandType = adCmdText
                 .CommandText = strconsulta
                 Set parametro = .CreateParameter(, adVarChar, adParamInput, 100, var_codigo)
                 .Parameters.Append parametro
            End With
            Set rsaux6 = comandoORA.execute
            Set comandoORA = Nothing
            Set parametro = Nothing
            
                        
            If Not rsaux6.EOF Then
               var_disponible = IIf(IsNull(rsaux6!Disponible), 0, rsaux6!Disponible)
            Else
               var_disponible = 0
            End If
            rsaux6.Close
            var_cantidad_pedida = IIf(IsNull(rsaux2!cantidad), 0, rsaux2!cantidad)
            If var_cantidad_pedida <= var_disponible Then
               var_Cantidad_posible = var_cantidad_pedida
            Else
               var_Cantidad_posible = var_disponible
            End If
            var_volumen_total = IIf(IsNull(rsaux2!cantidad), 0, rsaux2!cantidad) * var_volumen
            
            list_item.SubItems(2) = var_descripcion
            
            list_item.SubItems(3) = var_cantidad_pedida
            list_item.SubItems(4) = var_disponible
            list_item.SubItems(5) = var_Cantidad_posible
            list_item.SubItems(6) = var_Cantidad_posible * var_volumen
            list_item.SubItems(7) = var_volumen
            list_item.SubItems(8) = var_numero_archivo
            list_item.SubItems(9) = var_volumen_total
            Me.Refresh
            rsaux2.MoveNext
      Wend
      
   End If
   rsaux2.Close
   Me.frmbusqueda_pedido.Visible = False
   var_total_volumen = 0
   var_total_pedido = 0
   var_total_surtir = 0
   var_totla_volumne_pedido = 0
   For var_j = 1 To Me.lv_lista.ListItems.Count
       Me.lv_lista.ListItems.Item(var_j).Selected = True
       var_total_volumen = var_total_volumen + CDbl(Me.lv_lista.selectedItem.SubItems(6))
       var_total_pedido = var_total_pedido + CDbl(Me.lv_lista.selectedItem.SubItems(3))
       var_total_surtir = var_total_surtir + CDbl(Me.lv_lista.selectedItem.SubItems(5))
       var_totla_volumne_pedido = var_totla_volumne_pedido + CDbl(Me.lv_lista.selectedItem.SubItems(9))
   Next var_j
   Me.txt_volumen_pedidos = var_total_volumen
   If IsNumeric(Me.txt_volumen_unidad) Then
      Me.txt_porcentaje = (CDbl(Me.txt_volumen_pedidos) / CDbl(Me.txt_volumen_unidad)) * 100
   End If
   Me.txt_total_pedido = Format(var_total_pedido, "###,###,##0")
   Me.txt_total_surtir = Format(var_total_surtir, "###,###,##0")
   Me.txt_volumen_total = Format(var_totla_volumne_pedido, "###,###,##0")
   
   Call ilumina_grid
   Exit Sub
SALIR:
   MsgBox "Surgio un error al cargar el archivo, debe de tener las columnas CODIGO Y CANTIDAD y la hoja debe ded llamarse PLANTILLA PEDIDO", vbOKOnly, "ATENCION"
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
End Sub

Private Sub cmd_cargar_archivo_Click()
   Me.frmbusqueda_pedido.Visible = True
   Me.Dir1.SetFocus
End Sub

Private Sub cmd_generar_archivos_Click()
   
   If Me.lv_lista.ListItems.Count > 0 Then
      var_si = MsgBox("¿Desea generar los archivos?", vbYesNo, "ATENCION")
      If var_si = 6 Then
         cnn.BeginTrans
         rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_GENERA_ARCHIVOS_PEDIDOS", cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value)
         Else
            var_consecutivo = 0
         End If
         rs.Close
         var_consecutivo = var_consecutivo + 1
         rs.Open "INSERT INTO TB_TEMP_ORACLE_GENERA_ARCHIVOS_PEDIDOS (INTE_TEM_CONSECUTIVO) VALUES (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
         cnn.CommitTrans
         For var_j = 1 To Me.lv_lista.ListItems.Count
             Me.lv_lista.ListItems.Item(var_j).Selected = True
             If CDbl(Me.lv_lista.selectedItem.SubItems(3)) > 0 Then
                var_cadena = "INSERT INTO TB_TEMP_ORACLE_GENERA_ARCHIVOS_PEDIDOS (INTE_TEM_CONSECUTIVO, ARCHIVO, CODIGO, CANTIDAD) VALUES ('" + CStr(var_consecutivo) + "','" + Me.lv_lista.selectedItem + "','" + Me.lv_lista.selectedItem.SubItems(1) + "'," + Me.lv_lista.selectedItem.SubItems(3) + ")"
                rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
             End If
         Next var_j
         rs.Open "SELECT DISTINCT ARCHIVO FROM TB_TEMP_ORACLE_GENERA_ARCHIVOS_PEDIDOS WHERE INTE_TEM_CONSECUTIVO = " + CStr(var_consecutivo) + " AND ARCHIVO IS NOT NULL", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set reporte = appl.OpenReport(App.Path + "\plantilla pedido.rpt")
               reporte.RecordSelectionFormula = "{TB_TEMP_ORACLE_GENERA_ARCHIVOS_PEDIDOS.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and {TB_TEMP_ORACLE_GENERA_ARCHIVOS_PEDIDOS.ARCHIVO} = '" + rs!archivo + "'"
               For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
               Next ntablas
               reporte.ExportOptions.FormatType = crEFTExcel80
               reporte.ExportOptions.DestinationType = crEDTDiskFile
               archivo = "c:\reportessid\GENERADO" + rs!archivo + ".xls"
               reporte.ExportOptions.DiskFileName = archivo
               reporte.Export False
               Set reporte = Nothing
               rs.MoveNext
         Wend
         rs.Close
         rs.Open "DELETE FROM TB_TEMP_ORACLE_GENERA_ARCHIVOS_PEDIDOS WHERE INTE_tEM_CONSECUTIVO = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         MsgBox "Se a terminado de generar los archivos", vbOKOnly, "ATENCION"
      End If
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Dir1_Change()
   Me.File1.Path = Me.Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub

Private Sub Drive1_Change()
   On Error GoTo SALIR:
   Me.Dir1.Path = Me.Drive1.Drive
   Me.Dir1.Refresh
   Exit Sub
SALIR:
   MsgBox "Unidad incorrecta"
   Me.Drive1.Drive = "c:"
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub

Private Sub File1_Click()
   If CStr(Me.Dir1.Path) = "C:\" Or CStr(Me.Dir1.Path) = "c:\" Then
      Me.txt_ruta = CStr(Me.Dir1.Path) + Me.File1.FileName
   Else
      Me.txt_ruta = CStr(Me.Dir1.Path) + "\" + Me.File1.FileName
   End If
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
   Me.frm_cantidad_eliminar.Visible = False
   Me.frm_cantidad_agregar.Visible = False
   Me.frmbusqueda_pedido.Visible = False
   'rs.Open "select * from tb_oracle_tipo_unidades", cnn, adOpenDynamic, adLockBatchOptimistic
   rs.Open "SELECT NOMBRE, VOLUMEN  FROM tb_oracle_transportes WHERE EXPORTACIONES = 1 order by nombre", cnn, adOpenDynamic, adLockOptimistic
   Call RecsetToCombo(Me.cmb_tipo_unidades.hwnd, rs, 0)
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub



Private Sub lv_lista_KeyDown(KeyCode As Integer, Shift As Integer)
   'MsgBox CStr(KeyCode)
   If KeyCode = 46 Then
      var_renglon = Me.lv_lista.selectedItem.Index
      Me.lv_lista.selectedItem.SubItems(3) = 0
      Me.lv_lista.selectedItem.SubItems(5) = 0
      Me.lv_lista.selectedItem.SubItems(6) = 0
      var_total_volumen = 0
      var_total_pedido = 0
      var_total_surtir = 0
      For var_j = 1 To Me.lv_lista.ListItems.Count
          Me.lv_lista.ListItems.Item(var_j).Selected = True
          var_total_volumen = var_total_volumen + CDbl(Me.lv_lista.selectedItem.SubItems(6))
          var_total_pedido = var_total_pedido + CDbl(Me.lv_lista.selectedItem.SubItems(3))
          var_total_surtir = var_total_surtir + CDbl(Me.lv_lista.selectedItem.SubItems(5))
      Next var_j
      Me.txt_volumen_pedidos = var_total_volumen
      If IsNumeric(Me.txt_volumen_unidad) Then
         Me.txt_porcentaje = (CDbl(Me.txt_volumen_pedidos) / CDbl(Me.txt_volumen_unidad)) * 100
      End If
      Me.txt_total_pedido = Format(var_total_pedido, "###,###,##0")
      Me.txt_total_surtir = Format(var_total_surtir, "###,###,##0")
      Call ilumina_grid
      Me.lv_lista.ListItems(var_renglon).Selected = True
      Me.lv_lista.ListItems(var_renglon).EnsureVisible
      Me.lv_lista.SetFocus
   End If
   If KeyCode = 114 Then
      var_renglon = Me.lv_lista.selectedItem.Index
      Me.frm_cantidad_eliminar.Visible = True
      Me.txt_cantidad_eliminar = ""
      Me.txt_cantidad_eliminar.SetFocus
   End If
   If KeyCode = 115 Then
      var_renglon = Me.lv_lista.selectedItem.Index
      Me.frm_cantidad_agregar.Visible = True
      Me.txt_cantidad_agregar = ""
      Me.txt_cantidad_agregar.SetFocus
   End If
End Sub

Private Sub txt_cantidad_agregar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_agregar) Then
         If CDbl(Me.lv_lista.selectedItem.SubItems(4)) >= CDbl(Me.txt_cantidad_agregar) + CDbl(Me.lv_lista.selectedItem.SubItems(3)) Then
            Me.lv_lista.selectedItem.SubItems(3) = CDbl(Me.lv_lista.selectedItem.SubItems(3)) + CDbl(Me.txt_cantidad_agregar)
            Me.lv_lista.selectedItem.SubItems(5) = CDbl(Me.lv_lista.selectedItem.SubItems(5)) + CDbl(Me.txt_cantidad_agregar)
            Me.lv_lista.selectedItem.SubItems(6) = CDbl(Me.lv_lista.selectedItem.SubItems(5)) * CDbl(Me.lv_lista.selectedItem.SubItems(7))
            var_total_volumen = 0
            var_total_pedido = 0
            var_total_surtir = 0
            For var_j = 1 To Me.lv_lista.ListItems.Count
                Me.lv_lista.ListItems.Item(var_j).Selected = True
                var_total_volumen = var_total_volumen + CDbl(Me.lv_lista.selectedItem.SubItems(6))
                var_total_pedido = var_total_pedido + CDbl(Me.lv_lista.selectedItem.SubItems(3))
                var_total_surtir = var_total_surtir + CDbl(Me.lv_lista.selectedItem.SubItems(5))
            Next var_j
            Me.txt_volumen_pedidos = var_total_volumen
            If IsNumeric(Me.txt_volumen_unidad) Then
               Me.txt_porcentaje = (CDbl(Me.txt_volumen_pedidos) / CDbl(Me.txt_volumen_unidad)) * 100
            End If
            Me.txt_total_pedido = Format(var_total_pedido, "###,###,##0")
            Me.txt_total_surtir = Format(var_total_surtir, "###,###,##0")
            Call ilumina_grid
            Me.lv_lista.ListItems(var_renglon).Selected = True
            Me.lv_lista.ListItems(var_renglon).EnsureVisible
            Me.lv_lista.SetFocus
         Else
            MsgBox "La cantidad pedida no debe ser mayor a la posible a surtir", vbOKOnly, "ATENCION"
            Me.frm_cantidad_agregar.Visible = False
         End If
      Else
         MsgBox "Cantidad a agregar incorrecta", vbOKOnly, "ATENCION"
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_cantidad_agregar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_agregar_LostFocus()
   Me.frm_cantidad_agregar.Visible = False
End Sub

Private Sub txt_cantidad_eliminar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsNumeric(Me.txt_cantidad_eliminar) Then
         If CDbl(Me.txt_cantidad_eliminar) <= CDbl(Me.lv_lista.selectedItem.SubItems(5)) Then
            Me.lv_lista.selectedItem.SubItems(5) = CDbl(Me.lv_lista.selectedItem.SubItems(5)) - CDbl(Me.txt_cantidad_eliminar)
            Me.lv_lista.selectedItem.SubItems(6) = CDbl(Me.lv_lista.selectedItem.SubItems(5)) * CDbl(Me.lv_lista.selectedItem.SubItems(7))
            var_total_volumen = 0
            var_total_pedido = 0
            var_total_surtir = 0
            For var_j = 1 To Me.lv_lista.ListItems.Count
                Me.lv_lista.ListItems.Item(var_j).Selected = True
                var_total_volumen = var_total_volumen + CDbl(Me.lv_lista.selectedItem.SubItems(6))
                var_total_pedido = var_total_pedido + CDbl(Me.lv_lista.selectedItem.SubItems(3))
                var_total_surtir = var_total_surtir + CDbl(Me.lv_lista.selectedItem.SubItems(5))
            Next var_j
            Me.txt_volumen_pedidos = var_total_volumen
            If IsNumeric(Me.txt_volumen_unidad) Then
               Me.txt_porcentaje = (CDbl(Me.txt_volumen_pedidos) / CDbl(Me.txt_volumen_unidad)) * 100
            End If
            Me.txt_total_pedido = Format(var_total_pedido, "###,###,##0")
            Me.txt_total_surtir = Format(var_total_surtir, "###,###,##0")
            Call ilumina_grid
            Me.lv_lista.ListItems(var_renglon).Selected = True
            Me.lv_lista.ListItems(var_renglon).EnsureVisible
            Me.lv_lista.SetFocus
         Else
            MsgBox "La cantidad a eliminar no puede ser mayor a la cantidad a surtir", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Cantidad a eliminar incorrecta", vbOKOnly, "ATENCION"
         Me.frm_cantidad_eliminar.Visible = False
      End If
   End If
   If KeyAscii = 27 Then
      Me.frm_cantidad_eliminar.Visible = False
   End If
End Sub

Private Sub txt_cantidad_eliminar_LostFocus()
   Me.frm_cantidad_eliminar.Visible = False
End Sub

Private Sub txt_porcentaje_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmb_tipo_unidades.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Me.frmbusqueda_pedido.Visible = False
   End If
End Sub

Private Sub txt_total_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
   Else
      Me.txt_total_surtir.SetFocus
   End If
End Sub

Private Sub txt_total_surtir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_total_pedido.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_volumen_pedidos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_porcentaje.SetFocus
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_volumen_unidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.txt_volumen_pedidos.SetFocus
   Else
      KeyAscii = 0
   End If
   
End Sub


Private Sub ilumina_grid()
    var_n = lv_lista.ListItems.Count
    For var_i = 1 To var_n
        lv_lista.ListItems.Item(var_i).Selected = True
        If Trim(lv_lista.selectedItem.SubItems(5)) = "0" Then
           lv_lista.ListItems.Item(var_i).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(1).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(2).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(3).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(4).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(5).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(6).Bold = False
           lv_lista.ListItems.Item(var_i).ForeColor = &HFF&
           lv_lista.ListItems.Item(var_i).ListSubItems(1).ForeColor = &HFF&
           lv_lista.ListItems.Item(var_i).ListSubItems(2).ForeColor = &HFF&
           lv_lista.ListItems.Item(var_i).ListSubItems(3).ForeColor = &HFF&
           lv_lista.ListItems.Item(var_i).ListSubItems(4).ForeColor = &HFF&
           lv_lista.ListItems.Item(var_i).ListSubItems(5).ForeColor = &HFF&
           lv_lista.ListItems.Item(var_i).ListSubItems(6).ForeColor = &HFF&
        Else
           lv_lista.ListItems.Item(var_i).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(1).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(2).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(3).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(4).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(5).Bold = False
           lv_lista.ListItems.Item(var_i).ListSubItems(6).Bold = False
           lv_lista.ListItems.Item(var_i).ForeColor = &H80000008
           lv_lista.ListItems.Item(var_i).ListSubItems(1).ForeColor = &H80000008
           lv_lista.ListItems.Item(var_i).ListSubItems(2).ForeColor = &H80000008
           lv_lista.ListItems.Item(var_i).ListSubItems(3).ForeColor = &H80000008
           lv_lista.ListItems.Item(var_i).ListSubItems(4).ForeColor = &H80000008
           lv_lista.ListItems.Item(var_i).ListSubItems(5).ForeColor = &H80000008
           lv_lista.ListItems.Item(var_i).ListSubItems(6).ForeColor = &H80000008
        End If
    Next var_i
    lv_lista.Refresh
End Sub

