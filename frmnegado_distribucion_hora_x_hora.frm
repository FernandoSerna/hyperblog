VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnegado_distribucion_hora_x_hora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tablero negado de distribución hora X hora"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20100
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   20100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3120
      Top             =   120
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      Picture         =   "frmnegado_distribucion_hora_x_hora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Actualizar "
      Top             =   83
      Width           =   375
   End
   Begin VB.TextBox txt_total_negado 
      Alignment       =   1  'Right Justify
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
      Left            =   17880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox txt_fecha 
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
      Left            =   960
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   20055
   End
   Begin MSComctlLib.ListView lv_pantalla 
      CausesValidation=   0   'False
      Height          =   4215
      Left            =   15
      TabIndex        =   1
      Top             =   600
      Width           =   20040
      _ExtentX        =   35348
      _ExtentY        =   7435
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pedido"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SKU"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripcion"
         Object.Width           =   7673
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Físico"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Apartado"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Disponible"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Negado"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Hora"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Ubicación 1"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Ubicación 2"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Causa Negado"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView lv_resumen 
      CausesValidation=   0   'False
      Height          =   4575
      Left            =   15
      TabIndex        =   2
      Top             =   5280
      Width           =   20040
      _ExtentX        =   35348
      _ExtentY        =   8070
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SKU"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   12700
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Físico"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Apartado"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Disponible"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Negado"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Hora"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Ubicación 1"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Ubicación 2"
         Object.Width           =   3881
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total negado:"
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
      Left            =   16200
      TabIndex        =   5
      Top             =   4860
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmnegado_distribucion_hora_x_hora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
      If IsDate(Me.txt_fecha) Then
         Me.lv_pantalla.ListItems.Clear
         Me.lv_resumen.ListItems.Clear
         Me.txt_total_negado = ""
         var_dia_anterior = CStr(Day(CDate(Me.txt_fecha) - 1))
         var_mes_anterior = CStr(Month(CDate(Me.txt_fecha) - 1))
         var_año_anterior = CStr(Year(CDate(Me.txt_fecha) - 1))
         If Len(var_dia_anterior) = 1 Then
            var_dia_anterior = "0" + var_dia_anterior
         End If
         If Len(var_mes_anterior) = 1 Then
            var_mes_anterior = "0" + var_mes_anterior
         End If
         If Len(var_año_anterior) = 1 Then
            var_año_anterior = "0" + var_año_anterior
         End If
         
         var_dia = CStr(Day(CDate(Me.txt_fecha)))
         var_mes = CStr(Month(CDate(Me.txt_fecha)))
         var_año = CStr(Year(CDate(Me.txt_fecha)))
         
         
         If Len(var_dia) = 1 Then
            var_dia = "0" + var_dia
         End If
         If Len(var_mes) = 1 Then
            var_mes = "0" + var_mes
         End If
         If Len(var_año) = 1 Then
            var_año = "0" + var_año
         End If
         
         
         var_fecha_anterior = var_dia_anterior + "/" + var_mes_anterior + "/" + var_año_anterior
         var_fecha = var_dia + "/" + var_mes + "/" + var_año
         x = 0
         rs.Open "alter session set nls_date_format = 'DD/MON/YYYY HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If x = 1 Then
         
         
         
         
         var_cadena = "select source_header_number, a.segment1 as codigo, b.description, fecha_negado, a.cantidad, a.CAUSA_NEGADO, nombre_causa_negado, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2"
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha_anterior + " 22:29:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 22:29:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " ORDER BY FECHA_NEGADO"
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_total = 0
         Me.lv_pantalla.ListItems.Clear
         While Not rsaux.EOF
               Set list_item = Me.lv_pantalla.ListItems.Add(, , rsaux!source_header_number)
               list_item.SubItems(1) = rsaux!codigo
               list_item.SubItems(2) = Format(rsaux!Description)
               list_item.SubItems(3) = Format(rsaux!CANTMANO)
               list_item.SubItems(4) = Format(rsaux!RESERVADA)
               list_item.SubItems(5) = Format(rsaux!Disponible)
               list_item.SubItems(6) = Format(rsaux!Cantidad)
               list_item.SubItems(7) = Format(rsaux!FECHA_NEGADO)
               list_item.SubItems(8) = Format(rsaux!ubicacion_1)
               list_item.SubItems(9) = Format(rsaux!ubicacion_2)
               list_item.SubItems(10) = rsaux!nombre_causa_negado
               var_total = var_total + rsaux!Cantidad
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.txt_total_negado = var_total
         Me.Refresh
         
         
         
         var_cadena = "select '22:29:00 a 22:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha_anterior + " 22:29:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha_anterior + " 22:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '22:29:00 a 22:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         Me.lv_resumen.ListItems.Clear
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         var_cadena = "select '23:00:00 a 23:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha_anterior + " 23:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha_anterior + " 23:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '23:00:00 a 23:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         var_cadena = "select '00:00:00 a 00:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 00:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 00:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '00:00:00 a 00:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
       
         var_cadena = "select '01:00:00 a 01:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 01:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 01:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '01:00:00 a 01:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         var_cadena = "select '02:00:00 a 02:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 02:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 02:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '02:00:00 a 02:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         var_cadena = "select '03:00:00 a 03:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 03:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 03:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '03:00:00 a 03:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
         var_cadena = "select '04:00:00 a 04:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 04:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 04:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '04:00:00 a 04:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
      
         var_cadena = "select '05:00:00 a 05:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 05:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 05:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '05:00:00 a 05:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '06:00:00 a 06:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 06:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 06:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '06:00:00 a 06:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
         var_cadena = "select '07:00:00 a 07:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 07:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 07:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '07:00:00 a 07:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
         var_cadena = "select '08:00:00 a 08:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 08:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 08:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '08:00:00 a 08:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
         var_cadena = "select '09:00:00 a 09:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 09:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 09:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '09:00:00 a 09:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '10:00:00 a 10:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 10:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 10:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '10:00:00 a 10:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
         var_cadena = "select '11:00:00 a 11:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 11:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 11:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '11:00:00 a 11:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
      
      
         var_cadena = "select '12:00:00 a 12:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 12:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 12:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '12:00:00 a 12:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
         var_cadena = "select '13:00:00 a 13:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 13:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 13:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '13:00:00 a 13:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
         var_cadena = "select '14:00:00 a 14:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 14:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 14:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '14:00:00 a 14:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         
         var_cadena = "select '15:00:00 a 15:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 15:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 15:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '15:00:00 a 15:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '16:00:00 a 16:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 16:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 16:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '16:00:00 a 16:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '17:00:00 a 17:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 17:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 17:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '17:00:00 a 17:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '18:00:00 a 18:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 18:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 18:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '18:00:00 a 18:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '19:00:00 a 19:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 19:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 19:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '19:00:00 a 19:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '20:00:00 a 20:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 20:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 20:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '20:00:00 a 20:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '21:00:00 a 21:59:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 21:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 21:59:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '21:00:00 a 21:59:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
      
         var_cadena = "select '22:00:00 a 22:29:59' as HORA, a.segment1 as codigo, b.description, SUM(a.cantidad) AS CANTIDAD, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2 "
         var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
         var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
         var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha + " 22:00:00','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 22:29:59','DD/MM/YYYY HH24:MI:SS')"
         var_cadena = var_cadena + " and a.segment1 =  b.segment1"
         var_cadena = var_cadena + " and a.organization_id = b.organization_id"
         var_cadena = var_cadena + " and c.organization_id = b.organization_id"
         var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
         var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
         var_cadena = var_cadena + " GROUP BY '22:00:00 a 22:29:59', a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2, b.attribute3"
      
         rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
         While Not rsaux.EOF
               Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
               list_item.SubItems(1) = Format(rsaux!Description)
               list_item.SubItems(2) = Format(rsaux!CANTMANO)
               list_item.SubItems(3) = Format(rsaux!RESERVADA)
               list_item.SubItems(4) = Format(rsaux!Disponible)
               list_item.SubItems(5) = Format(rsaux!Cantidad)
               list_item.SubItems(6) = Format(rsaux!hora)
               list_item.SubItems(7) = Format(rsaux!ubicacion_1)
               list_item.SubItems(8) = Format(rsaux!ubicacion_2)
               rsaux.MoveNext
         Wend
         rsaux.Close
         Me.Refresh
         Else
            
            var_cadena = "select source_header_number, a.segment1 as codigo, b.description, to_date(fecha_negado,'DD/MM/YYYY HH24:MI:SS') fecha_negado, a.cantidad, a.CAUSA_NEGADO, nombre_causa_negado, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2"
            var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
            var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
            var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha_anterior + " 22:29:00','DD/MM/YYYY HH24:MI:SS')"
            var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 22:29:59','DD/MM/YYYY HH24:MI:SS')"
            var_cadena = var_cadena + " and a.segment1 =  b.segment1"
            var_cadena = var_cadena + " and a.organization_id = b.organization_id"
            var_cadena = var_cadena + " and c.organization_id = b.organization_id"
            var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
            var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
            var_cadena = var_cadena + " ORDER BY FECHA_NEGADO desc"
            rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            var_total = 0
            Me.lv_pantalla.ListItems.Clear
            While Not rsaux.EOF
                  Set list_item = Me.lv_pantalla.ListItems.Add(, , rsaux!source_header_number)
                  list_item.SubItems(1) = rsaux!codigo
                  list_item.SubItems(2) = Format(rsaux!Description)
                  list_item.SubItems(3) = Format(rsaux!CANTMANO)
                  list_item.SubItems(4) = Format(rsaux!RESERVADA)
                  list_item.SubItems(5) = Format(rsaux!Disponible)
                  list_item.SubItems(6) = Format(rsaux!Cantidad)
                  list_item.SubItems(7) = Format(rsaux!FECHA_NEGADO)
                  list_item.SubItems(8) = Format(rsaux!ubicacion_1)
                  list_item.SubItems(9) = Format(rsaux!ubicacion_2)
                  list_item.SubItems(10) = rsaux!nombre_causa_negado
                  var_total = var_total + rsaux!Cantidad
                  rsaux.MoveNext
            Wend
            rsaux.Close
            For var_j = 1 To Me.lv_pantalla.ListItems.Count
                Me.lv_pantalla.ListItems(var_j).Selected = True
                Me.lv_pantalla.ListItems.Item(var_j).ListSubItems(6).Bold = True
            Next var_j
            
            Me.txt_total_negado = var_total
            Me.Refresh
            
            
            
            
            x = 1
            If x = 0 Then
            var_cadena = "select source_header_number, a.segment1 as codigo, b.description, fecha_negado, a.cantidad, a.CAUSA_NEGADO, nombre_causa_negado, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2,  substr(to_char(fecha_negado,'HH24:MI:SS'),1,2) HORA, substr(to_char(fecha_negado,'HH24:MI:SS'),4,2) minuto"
            var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
            var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
            var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha_anterior + " 22:29:00','DD/MM/YYYY HH24:MI:SS')"
            var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 22:29:59','DD/MM/YYYY HH24:MI:SS')"
            var_cadena = var_cadena + " and a.segment1 =  b.segment1"
            var_cadena = var_cadena + " and a.organization_id = b.organization_id"
            var_cadena = var_cadena + " and c.organization_id = b.organization_id"
            var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
            var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
            var_cadena = var_cadena + " ORDER BY FECHA_NEGADO desc"
            Text1 = var_cadena
            rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_hora = rsaux!hora
                  var_minuto = rsaux!minuto
            
                  If var_hora = 23 And var_minuto >= 30 Then
                     var_hora_minuto = "10:29:00 p.m. a 10:59:59 p.m."
                  Else
                     If var_hora = 0 Then
                        var_hora_minuto = "12:00:00 a.m. a 12:59:59 a.m."
                     End If
                     If var_hora = 1 Then
                        var_hora_minuto = "01:00:00 a.m. a 01:59:59 a.m."
                     End If
                     If var_hora = 2 Then
                        var_hora_minuto = "02:00:00 a.m. a 02:59:59 a.m."
                     End If
                     If var_hora = 3 Then
                        var_hora_minuto = "03:00:00 a.m. a 03:59:59 a.m."
                     End If
                     If var_hora = 4 Then
                        var_hora_minuto = "04:00:00 a.m. a 04:59:59 a.m."
                     End If
                     If var_hora = 5 Then
                        var_hora_minuto = "05:00:00 a.m. a 05:59:59 a.m."
                     End If
                     If var_hora = 6 Then
                        var_hora_minuto = "06:00:00 a.m. a 06:59:59 a.m."
                     End If
                     If var_hora = 7 Then
                        var_hora_minuto = "07:00:00 a.m. a 07:59:59 a.m."
                     End If
                     If var_hora = 8 Then
                        var_hora_minuto = "08:00:00 a.m. a 08:59:59 a.m."
                     End If
                     If var_hora = 9 Then
                        var_hora_minuto = "09:00:00 a.m. a 09:59:59 a.m."
                     End If
                     If var_hora = 10 Then
                        var_hora_minuto = "10:00:00 a.m. a 10:59:59 a.m."
                     End If
                     If var_hora = 11 Then
                        var_hora_minuto = "11:00:00 a.m. a 11:59:59 a.m."
                     End If
                     If var_hora = 12 Then
                        var_hora_minuto = "12:00:00 p.m. a 12:59:59 p.m."
                     End If
                     If var_hora = 13 Then
                        var_hora_minuto = "01:00:00 p.m. a 01:59:59 p.m."
                     End If
                     If var_hora = 14 Then
                        var_hora_minuto = "02:00:00 p.m. a 14:59:59 p.m."
                     End If
                     If var_hora = 15 Then
                        var_hora_minuto = "03:00:00 p.m. a 03:59:59 p.m."
                     End If
                     If var_hora = 16 Then
                        var_hora_minuto = "04:00:00 p.m. a 04:59:59 p.m."
                     End If
                     If var_hora = 17 Then
                        var_hora_minuto = "05:00:00 p.m. a 05:59:59 p.m."
                     End If
                     If var_hora = 18 Then
                        var_hora_minuto = "06:00:00 p.m. a 06:59:59 p.m."
                     End If
                     If var_hora = 19 Then
                        var_hora_minuto = "07:00:00 p.m. a 07:59:59 p.m."
                     End If
                     If var_hora = 20 Then
                        var_hora_minuto = "08:00:00 p.m. a 08:59:59 p.m."
                     End If
                     If var_hora = 21 Then
                        var_hora_minuto = "09:00:00 p.m. a 09:59:59 p.m."
                     End If
                     If var_hora = 22 Then
                        var_hora_minuto = "10:00:00 p.m. a 10:59:59 p.m."
                     End If
                     If var_hora = 23 Then
                        var_hora_minuto = "11:00:00 p.m. a 11:29:59 p.m."
                     End If
                  End If
            
                  Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
                  list_item.SubItems(1) = Format(rsaux!Description)
                  list_item.SubItems(2) = Format(rsaux!CANTMANO)
                  list_item.SubItems(3) = Format(rsaux!RESERVADA)
                  list_item.SubItems(4) = Format(rsaux!Disponible)
                  list_item.SubItems(5) = Format(rsaux!Cantidad)
                  list_item.SubItems(6) = var_hora_minuto
                  list_item.SubItems(7) = Format(rsaux!ubicacion_1)
                  list_item.SubItems(8) = Format(rsaux!ubicacion_2)
                  rsaux.MoveNext
            Wend
            rsaux.Close
            Else
            
            var_cadena = "select a.segment1 as codigo, b.description, sum(a.cantidad) as cantidad, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 as ubicacion_1, b.attribute3 ubicacion_2"
            var_cadena = var_cadena + " from xxvia_tb_negado_distribucion a, xxvia_system_items_b b, Xxvia_vw_existencias_inv c"
            var_cadena = var_cadena + " where nvl(causa_negado,' ') <>' '"
            var_cadena = var_cadena + " AND fecha_negado >= to_Date('" + var_fecha_anterior + " 22:29:00','DD/MM/YYYY HH24:MI:SS')"
            var_cadena = var_cadena + " AND FECHA_NEGADO < TO_DATE('" + var_fecha + " 22:29:59','DD/MM/YYYY HH24:MI:SS')"
            var_cadena = var_cadena + " and a.segment1 =  b.segment1"
            var_cadena = var_cadena + " and a.organization_id = b.organization_id"
            var_cadena = var_cadena + " and c.organization_id = b.organization_id"
            var_cadena = var_cadena + " and subinventory_code = 'CDI_ALMPT'"
            var_cadena = var_cadena + " and C.segment1 = B.SEGMENT1"
            var_cadena = var_cadena + " group by a.segment1, b.description, c.CANTMANO, c.RESERVADA, c.DISPONIBLE, b.attribute2 , b.attribute3"
            var_cadena = var_cadena + " ORDER BY cantidad desc, a.segment1 "
            Text1 = var_cadena
            rsaux.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
            While Not rsaux.EOF
                  var_hora = ""
                  var_minuto = ""
                  var_hora_minuto = ""
                  Set list_item = Me.lv_resumen.ListItems.Add(, , rsaux!codigo)
                  list_item.SubItems(1) = Format(rsaux!Description)
                  list_item.SubItems(2) = Format(rsaux!CANTMANO)
                  list_item.SubItems(3) = Format(rsaux!RESERVADA)
                  list_item.SubItems(4) = Format(rsaux!Disponible)
                  list_item.SubItems(5) = Format(rsaux!Cantidad)
                  list_item.SubItems(6) = var_hora_minuto
                  list_item.SubItems(7) = Format(rsaux!ubicacion_1)
                  list_item.SubItems(8) = Format(rsaux!ubicacion_2)
                  rsaux.MoveNext
            Wend
            rsaux.Close
            For var_j = 1 To Me.lv_resumen.ListItems.Count
                Me.lv_resumen.ListItems(var_j).Selected = True
                Me.lv_resumen.ListItems.Item(var_j).ListSubItems(5).Bold = True
            Next var_j
            
            End If
            Me.Refresh
         End If
      
      End If

End Sub

Private Sub Form_Load()
   Me.txt_fecha = Date
   Call Command4_Click
End Sub

Private Sub lv_pantalla_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_pantalla, ColumnHeader)
End Sub

Private Sub lv_resumen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(Me.lv_resumen, ColumnHeader)
   If ColumnHeader = "Negado" Then
      RgsT = lv_resumen.ListItems.Count
      For Ds = 1 To RgsT
          lv_resumen.ListItems.Item(Ds).SubItems(5) = FormatNumber _
          (lv_resumen.ListItems.Item(Ds).SubItems(5), 6)
      Next
      
      For Ds = 1 To RgsT
          If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 10 Then
             lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(1, "0") & _
             lv_resumen.ListItems.Item(Ds).SubItems(5)
          End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 9 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(2, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 8 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(3, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 7 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(4, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 6 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(5, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 5 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(6, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 4 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(5, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 3 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(4, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 2 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(3, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 1 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(2, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
         If Len(lv_resumen.ListItems.Item(Ds).SubItems(5)) = 0 Then
            lv_resumen.ListItems.Item(Ds).SubItems(5) = String$(1, "0") & _
            lv_resumen.ListItems.Item(Ds).SubItems(5)
         End If
      Next
      If lv_resumen.SortOrder = lvwAscending Then
         lv_resumen.SortKey = 5
         lv_resumen.SortOrder = lvwDescending
      Else
         lv_resumen.SortKey = 5
         lv_resumen.SortOrder = lvwAscending
      End If
      For Ds = 1 To RgsT
          lv_resumen.ListItems.Item(Ds).SubItems(5) = FormatNumber _
          (lv_resumen.ListItems.Item(Ds).SubItems(5), 0)
      Next
     ' Call pro_ordena_listas(Me.lv_resumen, ColumnHeader)
   Else
      Call pro_ordena_listas(Me.lv_resumen, ColumnHeader)
   End If
End Sub

Private Sub Timer1_Timer()
   Call Command4_Click
End Sub

Private Sub txt_fecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Command4_Click
   End If
End Sub

