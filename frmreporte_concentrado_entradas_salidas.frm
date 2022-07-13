VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmreporte_concentrado_entradas_salidas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Concentrado de Entradas y Salidas"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4485
   Begin VB.Frame frm_lista 
      Height          =   2085
      Left            =   0
      TabIndex        =   13
      Top             =   480
      Width           =   4485
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1635
         Left            =   45
         TabIndex        =   14
         Top             =   405
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   2884
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5380
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   15
         Top             =   120
         Width           =   4410
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Reporte "
      Height          =   945
      Left            =   75
      TabIndex        =   16
      Top             =   1875
      Width           =   4350
      Begin VB.OptionButton opt_precio 
         Caption         =   "Precio"
         Height          =   270
         Left            =   150
         TabIndex        =   8
         Top             =   600
         Width           =   1665
      End
      Begin VB.OptionButton opt_costo 
         Caption         =   "Costo"
         Height          =   270
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Almacen "
      Height          =   645
      Left            =   60
      TabIndex        =   12
      Top             =   450
      Width           =   4380
      Begin VB.TextBox txt_nombre_almacen 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   210
         Width           =   3375
      End
      Begin VB.TextBox txt_clave_almacen 
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   210
         Width           =   870
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   75
      TabIndex        =   9
      Top             =   1155
      Width           =   4335
      Begin VB.TextBox txt_inicio 
         Height          =   315
         Left            =   690
         TabIndex        =   5
         Top             =   255
         Width           =   1080
      End
      Begin VB.TextBox txt_fin 
         Height          =   315
         Left            =   2700
         TabIndex        =   6
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2385
         TabIndex        =   10
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   345
      Width           =   4485
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      Picture         =   "frmreporte_concentrado_entradas_salidas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4095
      Picture         =   "frmreporte_concentrado_entradas_salidas.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
End
Attribute VB_Name = "frmreporte_concentrado_entradas_salidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xlApp1 As Excel.Application
Dim var_contador As Integer
Dim var_i As Integer, var_j As Integer

Private Sub reporte_precio()
   Set xlApp1 = New Excel.Application
   Dim xlApp2 As Object
   Set xlApp2 = CreateObject("Excel.Application")
   xlApp1.Visible = True
   xlApp1.Visible = True
   xlApp1.Workbooks.Add
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(2))).Value = "CONCENTRADO DE ENTRADAS Y SALIDAS A PRECIO"
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(2))).Font.Bold = True
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(3))).Value = rs!vcha_alm_nombre
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(3))).Font.Bold = True
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(4))).Value = "Del " + CStr(rs!dtim_tem_fecha_inicio) + " al " + CStr(rs!dtim_tem_fecha_fin)
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(4))).Font.Bold = True
   var_i = 6
   xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i))).Value = "LINEA"
   xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i))).Font.Bold = True
   xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i))).Value = "INVENTARIO INICIAL"
   xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i))).Font.Bold = True
   xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i)) + ":D" + Trim(CStr(var_i))).Merge
   var_contador = 0
   var_j = 1
   While var_j < 20
         If var_j = 1 Then
            If Not IsNull(rs!vcha_tem_movimiento_1) Then
               xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_1
               xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i)) + ":G" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_2) Then
               xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = rs!vcha_tem_nombre_movimiento_2
               xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i)) + ":J" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_3) Then
               xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_3
               xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i)) + ":M" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_4) Then
               xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_4
               xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i)) + ":P" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_5) Then
               xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_5
               xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i)) + ":S" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_6) Then
               xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_6
               xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i)) + ":V" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_7) Then
               xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_7
               xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i)) + ":Y" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_8) Then
               xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_8
               xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i)) + ":AB" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_9) Then
               xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_9
               xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i)) + ":AE" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_10) Then
               xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_10
               xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i)) + ":AH" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_11) Then
               xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_11
               xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i)) + ":AK" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_12) Then
               xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_12
               xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i)) + ":AN" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_13) Then
               xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_13
               xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i)) + ":AQ" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_14) Then
               xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_14
               xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i)) + ":AT" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_15) Then
               xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_15
               xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i)) + ":AW" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_16) Then
               xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_16
               xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i)) + ":AZ" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_17) Then
               xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_17
               xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i)) + ":BC" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_18) Then
               xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_18
               xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i)) + ":BF" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_19) Then
               xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_19
               xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i)) + ":BI" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_20) Then
               xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_20
               xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i)) + ":BL" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
         End If
         var_j = var_j + 1
   Wend
   If var_contador = 0 Then
      xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i)) + ":F" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 1 Then
      xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i)) + ":I" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 2 Then
      xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i)) + ":L" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 3 Then
      xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i)) + ":O" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 4 Then
      xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i)) + ":R" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 5 Then
      xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i)) + ":U" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 6 Then
      xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i)) + ":X" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 7 Then
      xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i)) + ":AA" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 8 Then
      xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i)) + ":AD" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 9 Then
      xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i)) + ":AG" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 10 Then
      xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i)) + ":AJ" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 11 Then
      xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i)) + ":A,M" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 12 Then
      xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i)) + ":AP" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 13 Then
      xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i)) + ":AS" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 14 Then
      xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i)) + ":AV" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 15 Then
      xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i)) + ":AY" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 16 Then
      xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i)) + ":BB" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 17 Then
      xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i)) + ":BE" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 18 Then
      xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i)) + ":BH" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 19 Then
      xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i)) + ":BK" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 17 Then
      xlApp1.Worksheets(1).Range("BM" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("BM" + Trim(CStr(var_i)) + ":BN" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("BM" + Trim(CStr(var_i))).Font.Bold = True
   End If
   var_i = var_i + 1
   
   xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i))).Value = ""
   xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i))).Value = "CANTIDAD"
   xlApp1.Worksheets(1).Range("D" + Trim(CStr(var_i))).Value = "PRECIO"
   For var_j = 1 To var_contador
       If var_j = 1 Then
          xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("F" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 2 Then
          xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("J" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 3 Then
          xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("L" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 4 Then
          xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("P" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 5 Then
          xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("R" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 6 Then
          xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("V" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 7 Then
          xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("X" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 8 Then
          xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AB" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 9 Then
          xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AD" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 10 Then
          xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AH" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 11 Then
          xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AJ" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 12 Then
          xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AN" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 13 Then
          xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AP" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 14 Then
          xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AS" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AT" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 15 Then
          xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AV" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AW" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 16 Then
          xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AY" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("AZ" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 17 Then
          xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("BB" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("BC" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 18 Then
          xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("BE" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("BF" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 19 Then
          xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("BH" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("BI" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
       If var_j = 20 Then
          xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("BK" + Trim(CStr(var_i))).Value = "PRECIO"
          xlApp1.Worksheets(1).Range("BL" + Trim(CStr(var_i))).Value = "DESCUENTO"
       End If
    Next var_j
    
   If var_contador = 0 Then
      xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("F" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 1 Then
      xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 2 Then
      xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("L" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 3 Then
      xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 4 Then
      xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("R" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 5 Then
      xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 6 Then
      xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("X" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 7 Then
      xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 8 Then
      xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AD" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 9 Then
      xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 10 Then
      xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AJ" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 11 Then
      xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 12 Then
      xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AP" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 13 Then
      xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AS" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 14 Then
      xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AV" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 15 Then
      xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AY" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 16 Then
      xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("BB" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 17 Then
      xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("BE" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 18 Then
      xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("BH" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 19 Then
      xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("BK" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   If var_contador = 20 Then
      xlApp1.Worksheets(1).Range("BM" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("BN" + Trim(CStr(var_i))).Value = "PRECIO"
   End If
   xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i)) + ":BN" + Trim(CStr(var_i))).Font.Bold = True
    
    
    
    
    
    var_i = var_i + 1
    While Not rs.EOF
          xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!vcha_lin_nombre), "", rs!vcha_lin_nombre)
          xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_INICIAL), "", rs!FLOA_TEM_INVENTARIO_INICIAL)
          xlApp1.Worksheets(1).Range("D" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_INICIAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_INICIAL_PRECIO)
          
          xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_1), "", rs!FLOA_TEM_CANTIDAD_1)
          xlApp1.Worksheets(1).Range("F" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_1), "", rs!FLOA_TEM_PRECIO_1)
          xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_1), "", rs!FLOA_TEM_DESCUENTO_1)
          
          xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_2), "", rs!FLOA_TEM_CANTIDAD_2)
          xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_2), "", rs!FLOA_TEM_PRECIO_2)
          xlApp1.Worksheets(1).Range("J" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_2), "", rs!FLOA_TEM_DESCUENTO_2)
          
          xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_3), "", rs!FLOA_TEM_CANTIDAD_3)
          xlApp1.Worksheets(1).Range("L" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_3), "", rs!FLOA_TEM_PRECIO_3)
          xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_3), "", rs!FLOA_TEM_DESCUENTO_3)
          
          xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_4), "", rs!FLOA_TEM_CANTIDAD_4)
          xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_4), "", rs!FLOA_TEM_PRECIO_4)
          xlApp1.Worksheets(1).Range("P" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_4), "", rs!FLOA_TEM_DESCUENTO_4)
          
          xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_5), "", rs!FLOA_TEM_CANTIDAD_5)
          xlApp1.Worksheets(1).Range("R" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_5), "", rs!FLOA_TEM_PRECIO_5)
          xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_5), "", rs!FLOA_TEM_DESCUENTO_5)
          
          xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_6), "", rs!FLOA_TEM_CANTIDAD_6)
          xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_6), "", rs!FLOA_TEM_PRECIO_6)
          xlApp1.Worksheets(1).Range("V" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_6), "", rs!FLOA_TEM_DESCUENTO_6)
          
          xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_7), "", rs!FLOA_TEM_CANTIDAD_7)
          xlApp1.Worksheets(1).Range("X" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_7), "", rs!FLOA_TEM_PRECIO_7)
          xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_7), "", rs!FLOA_TEM_DESCUENTO_7)
          
          xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_8), "", rs!FLOA_TEM_CANTIDAD_8)
          xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_8), "", rs!FLOA_TEM_PRECIO_8)
          xlApp1.Worksheets(1).Range("AB" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_8), "", rs!FLOA_TEM_DESCUENTO_8)
          
          xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_9), "", rs!FLOA_TEM_CANTIDAD_9)
          xlApp1.Worksheets(1).Range("AD" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_9), "", rs!FLOA_TEM_PRECIO_9)
          xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_9), "", rs!FLOA_TEM_DESCUENTO_9)
          
          xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_10), "", rs!FLOA_TEM_CANTIDAD_10)
          xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_10), "", rs!FLOA_TEM_PRECIO_10)
          xlApp1.Worksheets(1).Range("AH" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_10), "", rs!FLOA_TEM_DESCUENTO_10)
          
          xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_11), "", rs!FLOA_TEM_CANTIDAD_11)
          xlApp1.Worksheets(1).Range("AJ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_11), "", rs!FLOA_TEM_PRECIO_11)
          xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_11), "", rs!FLOA_TEM_DESCUENTO_11)
          
          xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_12), "", rs!FLOA_TEM_CANTIDAD_12)
          xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_12), "", rs!FLOA_TEM_PRECIO_12)
          xlApp1.Worksheets(1).Range("AN" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_12), "", rs!FLOA_TEM_DESCUENTO_12)
          
          xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_13), "", rs!FLOA_TEM_CANTIDAD_13)
          xlApp1.Worksheets(1).Range("AP" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_13), "", rs!FLOA_TEM_PRECIO_13)
          xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_13), "", rs!FLOA_TEM_DESCUENTO_13)
          
          xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_14), "", rs!FLOA_TEM_CANTIDAD_14)
          xlApp1.Worksheets(1).Range("AS" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_14), "", rs!FLOA_TEM_PRECIO_14)
          xlApp1.Worksheets(1).Range("AT" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_14), "", rs!FLOA_TEM_DESCUENTO_14)
          
          xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_15), "", rs!FLOA_TEM_CANTIDAD_15)
          xlApp1.Worksheets(1).Range("AV" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_15), "", rs!FLOA_TEM_PRECIO_15)
          xlApp1.Worksheets(1).Range("AW" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_15), "", rs!FLOA_TEM_DESCUENTO_15)
          
          xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_16), "", rs!FLOA_TEM_CANTIDAD_16)
          xlApp1.Worksheets(1).Range("AY" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_16), "", rs!FLOA_TEM_PRECIO_16)
          xlApp1.Worksheets(1).Range("AZ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_16), "", rs!FLOA_TEM_DESCUENTO_16)
          
          xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_17), "", rs!FLOA_TEM_CANTIDAD_17)
          xlApp1.Worksheets(1).Range("BB" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_17), "", rs!FLOA_TEM_PRECIO_17)
          xlApp1.Worksheets(1).Range("BC" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_17), "", rs!FLOA_TEM_DESCUENTO_17)
          
          xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_18), "", rs!FLOA_TEM_CANTIDAD_18)
          xlApp1.Worksheets(1).Range("BE" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_18), "", rs!FLOA_TEM_PRECIO_18)
          xlApp1.Worksheets(1).Range("BF" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_19), "", rs!FLOA_TEM_DESCUENTO_19)
          
          xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_19), "", rs!FLOA_TEM_CANTIDAD_19)
          xlApp1.Worksheets(1).Range("BH" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_19), "", rs!FLOA_TEM_PRECIO_19)
          xlApp1.Worksheets(1).Range("BI" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_19), "", rs!FLOA_TEM_DESCUENTO_19)
          
          xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_20), "", rs!FLOA_TEM_CANTIDAD_20)
          xlApp1.Worksheets(1).Range("BK" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_PRECIO_20), "", rs!FLOA_TEM_PRECIO_20)
          xlApp1.Worksheets(1).Range("BL" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_DESCUENTO_19), "", rs!FLOA_TEM_DESCUENTO_19)
          
          
          If var_contador = 0 Then
             xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("F" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 1 Then
             xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 2 Then
             xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("L" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 3 Then
             xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 4 Then
             xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("R" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 5 Then
             xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 6 Then
             xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("X" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 7 Then
             xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 8 Then
             xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AD" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 9 Then
             xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 10 Then
             xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AJ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 11 Then
             xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 12 Then
             xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AP" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 13 Then
             xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AS" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 14 Then
             xlApp1.Worksheets(1).Range("AU" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AV" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 15 Then
             xlApp1.Worksheets(1).Range("AX" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("AY" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 16 Then
             xlApp1.Worksheets(1).Range("BA" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("BB" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 17 Then
             xlApp1.Worksheets(1).Range("BD" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("BE" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 18 Then
             xlApp1.Worksheets(1).Range("BG" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("BH" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 19 Then
             xlApp1.Worksheets(1).Range("BJ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("BK" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          If var_contador = 20 Then
             xlApp1.Worksheets(1).Range("BM" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
             xlApp1.Worksheets(1).Range("BN" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO), "", rs!FLOA_TEM_INVENTARIO_FINAL_PRECIO)
          End If
          
          
          xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i)) + ":BN" + Trim(CStr(var_i))).NumberFormat = "###,###,##0.00"
          var_i = var_i + 1
          rs.MoveNext
    Wend
    Call totales_precios
     xlApp1.Worksheets(1).Range("C" + Trim(Str(var_i)) + ":BL" + Trim(Str(var_i))).Font.Bold = True
End Sub
Private Sub totales_precios()
    xlApp1.Worksheets(1).Range("B" + Trim(Str(var_i))).Formula = "TOTALES"
    xlApp1.Worksheets(1).Range("B" + Trim(Str(var_i))).Font.Bold = True
    xlApp1.Worksheets(1).Range("C" + Trim(Str(var_i))).Formula = "=sum(C7:C" + Trim(CStr(var_i - 1)) + ")"
    xlApp1.Worksheets(1).Range("C" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    xlApp1.Worksheets(1).Range("D" + Trim(Str(var_i))).Formula = "=sum(D7:D" + Trim(CStr(var_i - 1)) + ")"
    xlApp1.Worksheets(1).Range("D" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    For var_j = 1 To var_contador
        If var_j = 1 Then
           xlApp1.Worksheets(1).Range("E" + Trim(Str(var_i))).Formula = "=sum(E7:E" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("E" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("F" + Trim(Str(var_i))).Formula = "=sum(F7:F" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("F" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("G" + Trim(Str(var_i))).Formula = "=sum(G7:G" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("G" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 2 Then
           xlApp1.Worksheets(1).Range("H" + Trim(Str(var_i))).Formula = "=sum(H7:H" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("H" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("I" + Trim(Str(var_i))).Formula = "=sum(I7:I" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("I" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("J" + Trim(Str(var_i))).Formula = "=sum(J7:J" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("J" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        'AQUI
        If var_j = 3 Then
           xlApp1.Worksheets(1).Range("K" + Trim(Str(var_i))).Formula = "=sum(K7:K" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("K" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("L" + Trim(Str(var_i))).Formula = "=sum(L7:L" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("L" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("M" + Trim(Str(var_i))).Formula = "=sum(M7:M" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("M" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 4 Then
           xlApp1.Worksheets(1).Range("N" + Trim(Str(var_i))).Formula = "=sum(N7:N" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("N" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("O" + Trim(Str(var_i))).Formula = "=sum(O7:O" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("O" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("P" + Trim(Str(var_i))).Formula = "=sum(P7:P" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("P" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 5 Then
           xlApp1.Worksheets(1).Range("Q" + Trim(Str(var_i))).Formula = "=sum(Q7:Q" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("Q" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("R" + Trim(Str(var_i))).Formula = "=sum(R7:R" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("R" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("S" + Trim(Str(var_i))).Formula = "=sum(S7:S" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("S" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 6 Then
           xlApp1.Worksheets(1).Range("T" + Trim(Str(var_i))).Formula = "=sum(T7:T" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("T" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("U" + Trim(Str(var_i))).Formula = "=sum(U7:U" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("U" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("V" + Trim(Str(var_i))).Formula = "=sum(V7:V" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("V" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 7 Then
           xlApp1.Worksheets(1).Range("W" + Trim(Str(var_i))).Formula = "=sum(W7:W" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("W" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("X" + Trim(Str(var_i))).Formula = "=sum(X7:X" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("X" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("Y" + Trim(Str(var_i))).Formula = "=sum(Y7:Y" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("Y" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 8 Then
           xlApp1.Worksheets(1).Range("Z" + Trim(Str(var_i))).Formula = "=sum(Z7:Z" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("Z" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AA" + Trim(Str(var_i))).Formula = "=sum(AA7:AA" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AA" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AB" + Trim(Str(var_i))).Formula = "=sum(AB7:AB" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AB" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 9 Then
           xlApp1.Worksheets(1).Range("AC" + Trim(Str(var_i))).Formula = "=sum(AC7:AC" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AC" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AD" + Trim(Str(var_i))).Formula = "=sum(AD7:AD" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AD" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AE" + Trim(Str(var_i))).Formula = "=sum(AE7:AE" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AE" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 10 Then
           xlApp1.Worksheets(1).Range("AF" + Trim(Str(var_i))).Formula = "=sum(AF7:AF" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AF" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AG" + Trim(Str(var_i))).Formula = "=sum(AG7:AG" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AG" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AH" + Trim(Str(var_i))).Formula = "=sum(AH7:AH" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AH" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 11 Then
           xlApp1.Worksheets(1).Range("AI" + Trim(Str(var_i))).Formula = "=sum(AI7:AI" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AI" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AJ" + Trim(Str(var_i))).Formula = "=sum(AJ7:AJ" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AJ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AK" + Trim(Str(var_i))).Formula = "=sum(AK7:AK" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AK" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 12 Then
           xlApp1.Worksheets(1).Range("AL" + Trim(Str(var_i))).Formula = "=sum(AL7:AL" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AL" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AM" + Trim(Str(var_i))).Formula = "=sum(AM7:AM" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AM" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AN" + Trim(Str(var_i))).Formula = "=sum(AN7:AN" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AN" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 13 Then
           xlApp1.Worksheets(1).Range("AO" + Trim(Str(var_i))).Formula = "=sum(AO7:AO" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AO" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AP" + Trim(Str(var_i))).Formula = "=sum(AP7:AP" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AP" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AQ" + Trim(Str(var_i))).Formula = "=sum(AQ7:AQ" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AQ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        'AQUI
        If var_j = 14 Then
           xlApp1.Worksheets(1).Range("AR" + Trim(Str(var_i))).Formula = "=sum(AR7:AR" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AR" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AS" + Trim(Str(var_i))).Formula = "=sum(AS7:AS" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AS" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AT" + Trim(Str(var_i))).Formula = "=sum(AT7:AT" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AT" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 15 Then
           xlApp1.Worksheets(1).Range("AU" + Trim(Str(var_i))).Formula = "=sum(AU7:AU" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AU" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AV" + Trim(Str(var_i))).Formula = "=sum(AV7:AV" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AV" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AW" + Trim(Str(var_i))).Formula = "=sum(AW7:AW" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AW" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 16 Then
           xlApp1.Worksheets(1).Range("AX" + Trim(Str(var_i))).Formula = "=sum(AX7:AX" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AX" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AY" + Trim(Str(var_i))).Formula = "=sum(AY7:AY" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AY" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("AZ" + Trim(Str(var_i))).Formula = "=sum(AZ7:AZ" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("AZ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 17 Then
           xlApp1.Worksheets(1).Range("BA" + Trim(Str(var_i))).Formula = "=sum(BA7:BA" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BA" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("BB" + Trim(Str(var_i))).Formula = "=sum(BB7:BB" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BB" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("BC" + Trim(Str(var_i))).Formula = "=sum(BC7:BC" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BC" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 18 Then
           xlApp1.Worksheets(1).Range("BD" + Trim(Str(var_i))).Formula = "=sum(BD7:BD" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BD" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("BE" + Trim(Str(var_i))).Formula = "=sum(BE7:BE" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BE" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("BF" + Trim(Str(var_i))).Formula = "=sum(BF7:BF" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BF" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 19 Then
           xlApp1.Worksheets(1).Range("BG" + Trim(Str(var_i))).Formula = "=sum(BG7:BG" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BG" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("BH" + Trim(Str(var_i))).Formula = "=sum(BH7:BH" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BH" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("BI" + Trim(Str(var_i))).Formula = "=sum(BI7:BI" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BI" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
        If var_j = 20 Then
           xlApp1.Worksheets(1).Range("BJ" + Trim(Str(var_i))).Formula = "=sum(BJ7:BJ" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BJ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("BK" + Trim(Str(var_i))).Formula = "=sum(BK7:BK" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BK" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
           xlApp1.Worksheets(1).Range("BL" + Trim(Str(var_i))).Formula = "=sum(BL7:BL" + Trim(CStr(var_i - 1)) + ")"
           xlApp1.Worksheets(1).Range("BL" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
        End If
     Next var_j
     
    If var_contador = 0 Then
       xlApp1.Worksheets(1).Range("E" + Trim(Str(var_i))).Formula = "=sum(E7:E" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("E" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("F" + Trim(Str(var_i))).Formula = "=sum(F7:F" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("F" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 1 Then
       xlApp1.Worksheets(1).Range("H" + Trim(Str(var_i))).Formula = "=sum(H7:H" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("H" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("I" + Trim(Str(var_i))).Formula = "=sum(I7:I" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("I" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 2 Then
       xlApp1.Worksheets(1).Range("K" + Trim(Str(var_i))).Formula = "=sum(K7:K" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("K" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("L" + Trim(Str(var_i))).Formula = "=sum(L7:L" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("L" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 3 Then
       xlApp1.Worksheets(1).Range("N" + Trim(Str(var_i))).Formula = "=sum(N7:N" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("N" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("O" + Trim(Str(var_i))).Formula = "=sum(O7:O" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("O" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 4 Then
       xlApp1.Worksheets(1).Range("Q" + Trim(Str(var_i))).Formula = "=sum(Q7:Q" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("Q" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("R" + Trim(Str(var_i))).Formula = "=sum(R7:R" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("R" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 5 Then
       xlApp1.Worksheets(1).Range("T" + Trim(Str(var_i))).Formula = "=sum(T7:T" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("T" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("U" + Trim(Str(var_i))).Formula = "=sum(U7:U" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("U" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 6 Then
       xlApp1.Worksheets(1).Range("W" + Trim(Str(var_i))).Formula = "=sum(W7:W" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("W" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("X" + Trim(Str(var_i))).Formula = "=sum(X7:X" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("X" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 7 Then
       xlApp1.Worksheets(1).Range("Z" + Trim(Str(var_i))).Formula = "=sum(A7:A" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("Z" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AA" + Trim(Str(var_i))).Formula = "=sum(AA7:AA" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AA" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 8 Then
       xlApp1.Worksheets(1).Range("AC" + Trim(Str(var_i))).Formula = "=sum(AC7:AC" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AC" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AD" + Trim(Str(var_i))).Formula = "=sum(AD7:AD" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AD" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 9 Then
       xlApp1.Worksheets(1).Range("AF" + Trim(Str(var_i))).Formula = "=sum(AF7:AF" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AF" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AG" + Trim(Str(var_i))).Formula = "=sum(AG7:AG" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AG" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 10 Then
       xlApp1.Worksheets(1).Range("AI" + Trim(Str(var_i))).Formula = "=sum(AI7:AI" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AI" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AJ" + Trim(Str(var_i))).Formula = "=sum(AJ7:AJ" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AJ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 11 Then
       xlApp1.Worksheets(1).Range("AL" + Trim(Str(var_i))).Formula = "=sum(AL7:AL" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AL" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AM" + Trim(Str(var_i))).Formula = "=sum(AM7:AM" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AM" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 12 Then
       xlApp1.Worksheets(1).Range("AO" + Trim(Str(var_i))).Formula = "=sum(AO7:AO" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AO" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AP" + Trim(Str(var_i))).Formula = "=sum(AP7:AP" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AP" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 13 Then
       xlApp1.Worksheets(1).Range("AR" + Trim(Str(var_i))).Formula = "=sum(AR7:AR" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AR" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AS" + Trim(Str(var_i))).Formula = "=sum(AS7:AS" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AS" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 14 Then
       xlApp1.Worksheets(1).Range("AU" + Trim(Str(var_i))).Formula = "=sum(AU7:AU" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AU" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AV" + Trim(Str(var_i))).Formula = "=sum(AV7:AV" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AV" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 15 Then
       xlApp1.Worksheets(1).Range("AX" + Trim(Str(var_i))).Formula = "=sum(AX7:AX" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AX" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("AY" + Trim(Str(var_i))).Formula = "=sum(AY7:AY" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("AY" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 16 Then
       xlApp1.Worksheets(1).Range("BA" + Trim(Str(var_i))).Formula = "=sum(BA7:BA" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BA" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("BB" + Trim(Str(var_i))).Formula = "=sum(BB7:BB" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BB" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 17 Then
       xlApp1.Worksheets(1).Range("BD" + Trim(Str(var_i))).Formula = "=sum(BD7:BD" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BD" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("BE" + Trim(Str(var_i))).Formula = "=sum(BE7:BE" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BE" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 18 Then
       xlApp1.Worksheets(1).Range("BG" + Trim(Str(var_i))).Formula = "=sum(BG7:BG" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BG" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("BH" + Trim(Str(var_i))).Formula = "=sum(BH7:BH" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BH" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 19 Then
       xlApp1.Worksheets(1).Range("BJ" + Trim(Str(var_i))).Formula = "=sum(BJ7:BJ" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BJ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("BK" + Trim(Str(var_i))).Formula = "=sum(BK7:BK" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BK" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
    If var_contador = 20 Then
       xlApp1.Worksheets(1).Range("BM" + Trim(Str(var_i))).Formula = "=sum(BM7:BM" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BM" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       xlApp1.Worksheets(1).Range("BN" + Trim(Str(var_i))).Formula = "=sum(BN7:BN" + Trim(CStr(var_i - 1)) + ")"
       xlApp1.Worksheets(1).Range("BN" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
    End If
 
End Sub
Private Sub reporte_costo()
   Set xlApp1 = New Excel.Application
   Dim xlApp2 As Object
   Set xlApp2 = CreateObject("Excel.Application")
   xlApp1.Visible = True
   xlApp1.Visible = True
   xlApp1.Workbooks.Add
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(2))).Value = "CONCENTRADO DE ENTRADAS Y SALIDAS A COSTO"
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(2))).Font.Bold = True
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(3))).Value = rs!vcha_alm_nombre
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(3))).Font.Bold = True
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(4))).Value = "Del " + CStr(rs!dtim_tem_fecha_inicio) + " al " + CStr(rs!dtim_tem_fecha_fin)
   xlApp1.Worksheets(1).Range("A" + Trim(CStr(4))).Font.Bold = True
   var_i = 6
   xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i))).Value = "LINEA"
   xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i))).Font.Bold = True
   xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i))).Value = "INVENTARIO INICIAL"
   xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i))).Font.Bold = True
   xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i)) + ":D" + Trim(CStr(var_i))).Merge
   var_contador = 0
   var_j = 1
   While var_j < 20
         If var_j = 1 Then
            If Not IsNull(rs!vcha_tem_movimiento_1) Then
               xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_1
               xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i)) + ":F" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_2) Then
               xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Value = rs!vcha_tem_nombre_movimiento_2
               xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i)) + ":H" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_3) Then
               xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_3
               xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i)) + ":J" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_4) Then
               xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_4
               xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i)) + ":L" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_5) Then
               xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_5
               xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i)) + ":N" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_6) Then
               xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_6
               xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i)) + ":P" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_7) Then
               xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_7
               xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i)) + ":R" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_8) Then
               xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_8
               xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i)) + ":T" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_9) Then
               xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_9
               xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i)) + ":V" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_10) Then
               xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_10
               xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i)) + ":X" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_11) Then
               xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_11
               xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i)) + ":Z" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_12) Then
               xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_12
               xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i)) + ":AB" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_13) Then
               xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_13
               xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i)) + ":AD" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_14) Then
               xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_14
               xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i)) + ":AF" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_15) Then
               xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_15
               xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i)) + ":AH" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_16) Then
               xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_16
               xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i)) + ":AJ" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_17) Then
               xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_17
               xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i)) + ":AL" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_18) Then
               xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_18
               xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i)) + ":AN" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_19) Then
               xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_19
               xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i)) + ":AP" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
            If Not IsNull(rs!vcha_tem_movimiento_20) Then
               xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Value = rs!VCHA_TEM_NOMBRE_MOVIMIENTO_20
               xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i)) + ":AR" + Trim(CStr(var_i))).Merge
               xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Font.Bold = True
               var_contador = var_contador + 1
            Else
               var_j = 20
            End If
         End If
         var_j = var_j + 1
   Wend
   If var_contador = 0 Then
      xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i)) + ":F" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 1 Then
      xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i)) + ":H" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 2 Then
      xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i)) + ":J" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 3 Then
      xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i)) + ":L" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 4 Then
      xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i)) + ":N" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 5 Then
      xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i)) + ":P" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 6 Then
      xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i)) + ":R" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 7 Then
      xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i)) + ":T" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 8 Then
      xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i)) + ":V" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 9 Then
      xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i)) + ":X" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 10 Then
      xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i)) + ":Z" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 11 Then
      xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i)) + ":AB" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 12 Then
      xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i)) + ":AD" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 13 Then
      xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i)) + ":AF" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 14 Then
      xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i)) + ":AH" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 15 Then
      xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i)) + ":AJ" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 16 Then
      xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i)) + ":L" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 17 Then
      xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i)) + ":N" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 18 Then
      xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i)) + ":P" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 19 Then
      xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i)) + ":R" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Font.Bold = True
   End If
   If var_contador = 17 Then
      xlApp1.Worksheets(1).Range("AT" + Trim(CStr(var_i))).Value = "INVENTARIO FINAL"
      xlApp1.Worksheets(1).Range("AT" + Trim(CStr(var_i)) + ":U" + Trim(CStr(var_i))).Merge
      xlApp1.Worksheets(1).Range("AT" + Trim(CStr(var_i))).Font.Bold = True
   End If
   var_i = var_i + 1
   xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i))).Value = ""
   xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i))).Value = "CANTIDAD"
   xlApp1.Worksheets(1).Range("D" + Trim(CStr(var_i))).Value = "COSTO"
   For var_j = 1 To var_contador
       If var_j = 1 Then
          xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("F" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 2 Then
          xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 3 Then
          xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("J" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 4 Then
          xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("L" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 5 Then
          xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 6 Then
          xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("P" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 7 Then
          xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("R" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 8 Then
          xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 9 Then
          xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("V" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 10 Then
          xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("X" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 11 Then
          xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 12 Then
          xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AB" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 13 Then
          xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AD" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 14 Then
          xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 15 Then
          xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AH" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 16 Then
          xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AJ" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 17 Then
          xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 18 Then
          xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AN" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 19 Then
          xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AP" + Trim(CStr(var_i))).Value = "COSTO"
       End If
       If var_j = 20 Then
          xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Value = "CANTIDAD"
          xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = "COSTO"
       End If
   Next var_j
   If var_contador = 0 Then
      xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("F" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 1 Then
      xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 2 Then
      xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("J" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 3 Then
      xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("L" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 4 Then
      xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 5 Then
      xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("P" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 6 Then
      xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("R" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 7 Then
      xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 8 Then
      xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("V" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 9 Then
      xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("X" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 10 Then
      xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 11 Then
      xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AB" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 12 Then
      xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AD" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 13 Then
      xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 14 Then
      xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AH" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 15 Then
      xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AJ" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 16 Then
      xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 17 Then
      xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AN" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 18 Then
      xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AP" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 19 Then
      xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   If var_contador = 20 Then
      xlApp1.Worksheets(1).Range("AS" + Trim(CStr(var_i))).Value = "CANTIDAD"
      xlApp1.Worksheets(1).Range("AT" + Trim(CStr(var_i))).Value = "COSTO"
   End If
   xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i)) + ":AR" + Trim(CStr(var_i))).Font.Bold = True
   var_i = var_i + 1
   While Not rs.EOF
         xlApp1.Worksheets(1).Range("B" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!vcha_lin_nombre), "", rs!vcha_lin_nombre)
         xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_INICIAL), "", rs!FLOA_TEM_INVENTARIO_INICIAL)
         xlApp1.Worksheets(1).Range("D" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_INICIAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_INICIAL_COSTO)
         xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_1), "", rs!FLOA_TEM_CANTIDAD_1)
         xlApp1.Worksheets(1).Range("F" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_1), "", rs!FLOA_TEM_COSTO_1)
         xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_2), "", rs!FLOA_TEM_CANTIDAD_2)
         xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_2), "", rs!FLOA_TEM_COSTO_2)
         xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_3), "", rs!FLOA_TEM_CANTIDAD_3)
         xlApp1.Worksheets(1).Range("J" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_3), "", rs!FLOA_TEM_COSTO_3)
         xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_4), "", rs!FLOA_TEM_CANTIDAD_4)
         xlApp1.Worksheets(1).Range("L" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_4), "", rs!FLOA_TEM_COSTO_4)
         xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_5), "", rs!FLOA_TEM_CANTIDAD_5)
         xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_5), "", rs!FLOA_TEM_COSTO_5)
         xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_6), "", rs!FLOA_TEM_CANTIDAD_6)
         xlApp1.Worksheets(1).Range("P" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_6), "", rs!FLOA_TEM_COSTO_6)
         xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_7), "", rs!FLOA_TEM_CANTIDAD_7)
         xlApp1.Worksheets(1).Range("R" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_7), "", rs!FLOA_TEM_COSTO_7)
         xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_8), "", rs!FLOA_TEM_CANTIDAD_8)
         xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_8), "", rs!FLOA_TEM_COSTO_8)
         xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_9), "", rs!FLOA_TEM_CANTIDAD_9)
         xlApp1.Worksheets(1).Range("V" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_9), "", rs!FLOA_TEM_COSTO_9)
         xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_10), "", rs!FLOA_TEM_CANTIDAD_10)
         xlApp1.Worksheets(1).Range("X" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_10), "", rs!FLOA_TEM_COSTO_10)
         xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_11), "", rs!FLOA_TEM_CANTIDAD_11)
         xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_11), "", rs!FLOA_TEM_COSTO_11)
         xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_12), "", rs!FLOA_TEM_CANTIDAD_12)
         xlApp1.Worksheets(1).Range("AB" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_12), "", rs!FLOA_TEM_COSTO_12)
         xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_13), "", rs!FLOA_TEM_CANTIDAD_13)
         xlApp1.Worksheets(1).Range("AD" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_13), "", rs!FLOA_TEM_COSTO_13)
         xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_14), "", rs!FLOA_TEM_CANTIDAD_14)
         xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_14), "", rs!FLOA_TEM_COSTO_14)
         xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_15), "", rs!FLOA_TEM_CANTIDAD_15)
         xlApp1.Worksheets(1).Range("AH" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_15), "", rs!FLOA_TEM_COSTO_15)
         xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_16), "", rs!FLOA_TEM_CANTIDAD_16)
         xlApp1.Worksheets(1).Range("AJ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_16), "", rs!FLOA_TEM_COSTO_16)
         xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_17), "", rs!FLOA_TEM_CANTIDAD_17)
         xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_17), "", rs!FLOA_TEM_COSTO_17)
         xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_18), "", rs!FLOA_TEM_CANTIDAD_18)
         xlApp1.Worksheets(1).Range("AN" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_18), "", rs!FLOA_TEM_COSTO_18)
         xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_19), "", rs!FLOA_TEM_CANTIDAD_19)
         xlApp1.Worksheets(1).Range("AP" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_19), "", rs!FLOA_TEM_COSTO_19)
         xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_CANTIDAD_20), "", rs!FLOA_TEM_CANTIDAD_20)
         xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_COSTO_20), "", rs!FLOA_TEM_COSTO_20)
         xlApp1.Worksheets(1).Range("C" + Trim(CStr(var_i)) + ":AR" + Trim(CStr(var_i))).NumberFormat = "###,###,##0.00"
         If var_contador = 0 Then
            xlApp1.Worksheets(1).Range("E" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("F" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 1 Then
            xlApp1.Worksheets(1).Range("G" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("H" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 2 Then
            xlApp1.Worksheets(1).Range("I" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("J" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 3 Then
            xlApp1.Worksheets(1).Range("K" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("L" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 4 Then
            xlApp1.Worksheets(1).Range("M" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("N" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 5 Then
            xlApp1.Worksheets(1).Range("O" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("P" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 6 Then
            xlApp1.Worksheets(1).Range("Q" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("R" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 7 Then
            xlApp1.Worksheets(1).Range("S" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("T" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 8 Then
            xlApp1.Worksheets(1).Range("U" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("V" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 9 Then
            xlApp1.Worksheets(1).Range("W" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("X" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 10 Then
            xlApp1.Worksheets(1).Range("Y" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("Z" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 11 Then
            xlApp1.Worksheets(1).Range("AA" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AB" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 12 Then
            xlApp1.Worksheets(1).Range("AC" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AD" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 13 Then
            xlApp1.Worksheets(1).Range("AE" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AF" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 14 Then
            xlApp1.Worksheets(1).Range("AG" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AH" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 15 Then
            xlApp1.Worksheets(1).Range("AI" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AJ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 16 Then
            xlApp1.Worksheets(1).Range("AK" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AL" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 17 Then
            xlApp1.Worksheets(1).Range("AM" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AN" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 18 Then
            xlApp1.Worksheets(1).Range("AO" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AP" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 19 Then
            xlApp1.Worksheets(1).Range("AQ" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AR" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         If var_contador = 20 Then
            xlApp1.Worksheets(1).Range("AS" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD), "", rs!FLOA_TEM_INVENTARIO_FINAL_CANTIDAD)
            xlApp1.Worksheets(1).Range("AT" + Trim(CStr(var_i))).Value = IIf(IsNull(rs!FLOA_TEM_INVENTARIO_FINAL_COSTO), "", rs!FLOA_TEM_INVENTARIO_FINAL_COSTO)
         End If
         var_i = var_i + 1
         rs.MoveNext
   Wend
   xlApp1.Worksheets(1).Range("B" + Trim(Str(var_i))).Formula = "TOTALES"
   xlApp1.Worksheets(1).Range("B" + Trim(Str(var_i))).Font.Bold = True
   xlApp1.Worksheets(1).Range("C" + Trim(Str(var_i))).Formula = "=sum(C7:C" + Trim(CStr(var_i - 1)) + ")"
   xlApp1.Worksheets(1).Range("C" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   xlApp1.Worksheets(1).Range("D" + Trim(Str(var_i))).Formula = "=sum(D7:D" + Trim(CStr(var_i - 1)) + ")"
   xlApp1.Worksheets(1).Range("D" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   Call totales_excel
End Sub
Private Sub totales_excel()
   For var_j = 1 To var_contador
       If var_j = 1 Then
          xlApp1.Worksheets(1).Range("E" + Trim(Str(var_i))).Formula = "=sum(E7:E" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("E" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("F" + Trim(Str(var_i))).Formula = "=sum(F7:F" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("F" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 2 Then
          xlApp1.Worksheets(1).Range("G" + Trim(Str(var_i))).Formula = "=sum(G7:G" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("G" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("H" + Trim(Str(var_i))).Formula = "=sum(H7:H" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("H" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 3 Then
          xlApp1.Worksheets(1).Range("I" + Trim(Str(var_i))).Formula = "=sum(I7:I" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("I" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("J" + Trim(Str(var_i))).Formula = "=sum(J7:J" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("J" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 4 Then
          xlApp1.Worksheets(1).Range("K" + Trim(Str(var_i))).Formula = "=sum(K7:K" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("K" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("L" + Trim(Str(var_i))).Formula = "=sum(L7:L" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("L" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 5 Then
          xlApp1.Worksheets(1).Range("M" + Trim(Str(var_i))).Formula = "=sum(M7:M" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("M" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("N" + Trim(Str(var_i))).Formula = "=sum(N7:N" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("N" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 6 Then
          xlApp1.Worksheets(1).Range("O" + Trim(Str(var_i))).Formula = "=sum(O7:O" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("O" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("P" + Trim(Str(var_i))).Formula = "=sum(P7:P" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("P" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 7 Then
          xlApp1.Worksheets(1).Range("Q" + Trim(Str(var_i))).Formula = "=sum(Q7:Q" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("Q" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("R" + Trim(Str(var_i))).Formula = "=sum(R7:R" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("R" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 8 Then
          xlApp1.Worksheets(1).Range("S" + Trim(Str(var_i))).Formula = "=sum(S7:S" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("S" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("T" + Trim(Str(var_i))).Formula = "=sum(T7:T" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("T" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 9 Then
          xlApp1.Worksheets(1).Range("U" + Trim(Str(var_i))).Formula = "=sum(U7:U" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("U" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("V" + Trim(Str(var_i))).Formula = "=sum(V7:V" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("V" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 10 Then
          xlApp1.Worksheets(1).Range("W" + Trim(Str(var_i))).Formula = "=sum(W7:W" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("W" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("X" + Trim(Str(var_i))).Formula = "=sum(X7:X" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("X" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 11 Then
          xlApp1.Worksheets(1).Range("Y" + Trim(Str(var_i))).Formula = "=sum(Y7:Y" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("Y" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("Z" + Trim(Str(var_i))).Formula = "=sum(Z7:Z" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("Z" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 12 Then
          xlApp1.Worksheets(1).Range("AA" + Trim(Str(var_i))).Formula = "=sum(AA7:AA" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AA" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AB" + Trim(Str(var_i))).Formula = "=sum(AB7:AB" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AB" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 13 Then
          xlApp1.Worksheets(1).Range("AC" + Trim(Str(var_i))).Formula = "=sum(AC7:AC" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AC" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AD" + Trim(Str(var_i))).Formula = "=sum(AD7:AD" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AD" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 14 Then
          xlApp1.Worksheets(1).Range("AE" + Trim(Str(var_i))).Formula = "=sum(AE7:AE" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AE" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AF" + Trim(Str(var_i))).Formula = "=sum(AF7:AF" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AF" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 15 Then
          xlApp1.Worksheets(1).Range("AG" + Trim(Str(var_i))).Formula = "=sum(AG7:AG" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AG" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AH" + Trim(Str(var_i))).Formula = "=sum(AH7:AH" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AH" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 16 Then
          xlApp1.Worksheets(1).Range("AI" + Trim(Str(var_i))).Formula = "=sum(AI7:AI" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AI" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AJ" + Trim(Str(var_i))).Formula = "=sum(AI7:AI" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AJ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 17 Then
          xlApp1.Worksheets(1).Range("AK" + Trim(Str(var_i))).Formula = "=sum(AK7:AK" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AK" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AL" + Trim(Str(var_i))).Formula = "=sum(AL7:AL" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AL" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 18 Then
          xlApp1.Worksheets(1).Range("AM" + Trim(Str(var_i))).Formula = "=sum(AM7:AM" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AM" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AN" + Trim(Str(var_i))).Formula = "=sum(AN7:AN" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AN" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 19 Then
          xlApp1.Worksheets(1).Range("AO" + Trim(Str(var_i))).Formula = "=sum(AO7:AO" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AO" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AP" + Trim(Str(var_i))).Formula = "=sum(AP7:AP" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AP" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
       If var_j = 20 Then
          xlApp1.Worksheets(1).Range("AQ" + Trim(Str(var_i))).Formula = "=sum(AQ7:AQ" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AQ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
          xlApp1.Worksheets(1).Range("AR" + Trim(Str(var_i))).Formula = "=sum(AR7:AR" + Trim(CStr(var_i - 1)) + ")"
          xlApp1.Worksheets(1).Range("AR" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
       End If
   Next var_j
   If var_contador = 0 Then
      xlApp1.Worksheets(1).Range("E" + Trim(Str(var_i))).Formula = "=sum(E7:E" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("E" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("F" + Trim(Str(var_i))).Formula = "=sum(F7:F" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("F" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 1 Then
      xlApp1.Worksheets(1).Range("G" + Trim(Str(var_i))).Formula = "=sum(G7:G" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("G" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("H" + Trim(Str(var_i))).Formula = "=sum(H7:H" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("H" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 2 Then
      xlApp1.Worksheets(1).Range("I" + Trim(Str(var_i))).Formula = "=sum(I7:I" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("I" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("J" + Trim(Str(var_i))).Formula = "=sum(J7:J" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("J" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 3 Then
      xlApp1.Worksheets(1).Range("K" + Trim(Str(var_i))).Formula = "=sum(K7:K" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("K" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("L" + Trim(Str(var_i))).Formula = "=sum(L7:L" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("L" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 4 Then
      xlApp1.Worksheets(1).Range("M" + Trim(Str(var_i))).Formula = "=sum(M7:M" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("M" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("N" + Trim(Str(var_i))).Formula = "=sum(N7:N" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("N" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 5 Then
      xlApp1.Worksheets(1).Range("O" + Trim(Str(var_i))).Formula = "=sum(O7:O" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("O" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("P" + Trim(Str(var_i))).Formula = "=sum(P7:P" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("P" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 6 Then
      xlApp1.Worksheets(1).Range("Q" + Trim(Str(var_i))).Formula = "=sum(Q7:Q" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("Q" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("R" + Trim(Str(var_i))).Formula = "=sum(R7:R" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("R" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 7 Then
      xlApp1.Worksheets(1).Range("S" + Trim(Str(var_i))).Formula = "=sum(S7:S" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("S" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("T" + Trim(Str(var_i))).Formula = "=sum(T7:T" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("T" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 8 Then
      xlApp1.Worksheets(1).Range("U" + Trim(Str(var_i))).Formula = "=sum(U7:U" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("U" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("V" + Trim(Str(var_i))).Formula = "=sum(V7:V" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("V" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 9 Then
      xlApp1.Worksheets(1).Range("W" + Trim(Str(var_i))).Formula = "=sum(W7:W" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("W" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("X" + Trim(Str(var_i))).Formula = "=sum(X7:X" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("X" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 10 Then
      xlApp1.Worksheets(1).Range("Y" + Trim(Str(var_i))).Formula = "=sum(Y7:Y" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("Y" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("Z" + Trim(Str(var_i))).Formula = "=sum(Z7:Z" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("Z" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 11 Then
      xlApp1.Worksheets(1).Range("AA" + Trim(Str(var_i))).Formula = "=sum(AA7:AA" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AA" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AB" + Trim(Str(var_i))).Formula = "=sum(AB7:AB" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AB" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 12 Then
      xlApp1.Worksheets(1).Range("AC" + Trim(Str(var_i))).Formula = "=sum(AC7:AC" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AC" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AD" + Trim(Str(var_i))).Formula = "=sum(AD7:AD" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AD" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 13 Then
      xlApp1.Worksheets(1).Range("AE" + Trim(Str(var_i))).Formula = "=sum(AE7:AE" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AE" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AF" + Trim(Str(var_i))).Formula = "=sum(AF7:AF" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AF" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 14 Then
      xlApp1.Worksheets(1).Range("AG" + Trim(Str(var_i))).Formula = "=sum(AG7:AG" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AG" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AH" + Trim(Str(var_i))).Formula = "=sum(AH7:AH" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AH" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 15 Then
      xlApp1.Worksheets(1).Range("AI" + Trim(Str(var_i))).Formula = "=sum(AI7:AI" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AI" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AJ" + Trim(Str(var_i))).Formula = "=sum(AI7:AI" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AJ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 16 Then
      xlApp1.Worksheets(1).Range("AK" + Trim(Str(var_i))).Formula = "=sum(AK7:AK" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AK" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AL" + Trim(Str(var_i))).Formula = "=sum(AL7:AL" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AL" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 17 Then
      xlApp1.Worksheets(1).Range("AM" + Trim(Str(var_i))).Formula = "=sum(AM7:AM" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AM" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AN" + Trim(Str(var_i))).Formula = "=sum(AN7:AN" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AN" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 18 Then
      xlApp1.Worksheets(1).Range("AO" + Trim(Str(var_i))).Formula = "=sum(AO7:AO" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AO" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AP" + Trim(Str(var_i))).Formula = "=sum(AP7:AP" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AP" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 19 Then
      xlApp1.Worksheets(1).Range("AQ" + Trim(Str(var_i))).Formula = "=sum(AQ7:AQ" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AQ" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AR" + Trim(Str(var_i))).Formula = "=sum(AR7:AR" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AR" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   If var_contador = 20 Then
      xlApp1.Worksheets(1).Range("AS" + Trim(Str(var_i))).Formula = "=sum(AS7:AS" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AS" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
      xlApp1.Worksheets(1).Range("AT" + Trim(Str(var_i))).Formula = "=sum(AT7:AT" + Trim(CStr(var_i - 1)) + ")"
      xlApp1.Worksheets(1).Range("AT" + Trim(Str(var_i))).NumberFormat = "###,###,##0.00"
   End If
   xlApp1.Worksheets(1).Range("C" + Trim(Str(var_i)) + ":AR" + Trim(Str(var_i))).Font.Bold = True
End Sub
Private Sub cmd_imprimir_Click()
   Dim var_consecutivo As Integer, var_i As Integer
   Dim var_fecha_fin As String, var_fecha_inicio, var_fecha_fin_1
     
   If IsDate(txt_inicio) Then
      If IsDate(txt_fin) Then
         If CDate(txt_inicio) <= CDate(txt_fin) Then
            cnn.BeginTrans
            rs.Open "select max(INTE_TEM_CONSECUTIVO) as numero from TB_TEM_REPORTE_CONCENTRADO_ENTRADAS_SALIDAS", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               var_consecutivo = IIf(IsNull(rs!NUMERO), 0, rs!NUMERO)
            Else
               var_consecutivo = 0
            End If
            rs.Close
            var_consecutivo = var_consecutivo + 1
            rs.Open "insert into TB_TEM_REPORTE_CONCENTRADO_ENTRADAS_SALIDAS (inte_tem_consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            var_fecha_fin_1 = CDate(txt_fin) + 1
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
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            var_fecha_fin_1 = CDate(txt_fin)
            var_dia = CStr(Day(var_fecha_fin_1))
            var_mes = CStr(Month(var_fecha_fin_1))
            var_año = CStr(Year(var_fecha_fin_1))
            If Len(Trim(var_dia)) = 1 Then
               var_dia = "0" + var_dia
            End If
            If Len(Trim(var_mes)) = 1 Then
               var_mes = "0" + var_mes
            End If
            var_fecha_fin = "{d '" + var_año + "-" + var_mes + "-" + var_dia + "'}"
            
            rs.Open "execute SP_CONCENTRADO_ENTRADAS_SALIDAS " + CStr(var_consecutivo) + ", '" + txt_clave_almacen + "', " + var_fecha_inicio + ", " + var_fecha_fin + ", '" + var_clave_usuario_global + "', '" + fun_NombrePc + "'", cnn, adOpenDynamic, adLockOptimistic
            rs.Open "select * from VW_CONCENTRADO_ENTRADAS_SALIDAS_FINAL_LINEA where inte_tem_consecutivo = " + CStr(var_consecutivo) + " ORDER BY VCHA_LIN_NOMBRE", cnn, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
               If opt_costo = True Then
                  Call reporte_costo
               End If
               If opt_precio = True Then
                  Call reporte_precio
               End If
            Else
               MsgBox "No existe información para la consulta echa", vbOKOnly, "ATENCION"
            End If
            rs.Close
         Else
            MsgBox "La fecha de inicio debe de ser menor o igual a la fecha final", vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "Fecha final incorrecta", vbOKOnly, "ATENCION"
      End If
   Else
      MsgBox "Fecha de inicio incorrecta", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 2100
   Left = 3500
   txt_inicio = Date
   txt_fin = Date
   frm_lista.Visible = False
   opt_costo = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    var_swpassword = False
    var_modifica_registro = False
    Call activa_forma(var_activa_forma_reporte_concentrado_entradas_salidas)
End Sub

Private Sub lv_lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Call pro_ordena_listas(lv_lista, ColumnHeader)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If lv_lista.ListItems.Count > 0 Then
         txt_clave_almacen = lv_lista.selectedItem
         txt_nombre_almacen = lv_lista.selectedItem.SubItems(1)
      Else
         txt_clave_almacen = ""
         txt_nombre_almacen = ""
      End If
      txt_clave_almacen.SetFocus
      frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   frm_lista.Visible = False
End Sub

Private Sub txt_clave_almacen_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      rs.Open "select * from tb_almacenes where vcha_emp_empresa_id = '" + var_empresa + "' order by VCHA_ALM_NOMBRE", cnn, adOpenDynamic, adLockBatchOptimistic
      If Not rs.EOF Then
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!VCHA_ALM_ALMACEN_ID)
               list_item.SubItems(1) = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
               rs.MoveNext
         Wend
         rs.Close
         lbl_lista = "Almacenes"
         frm_lista.Visible = True
         lv_lista.SetFocus
      End If
   End If
End Sub

Private Sub txt_clave_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 39 Or KeyAscii = 61 Then
      KeyAscii = 0
   End If
   If KeyAscii = 13 Then
      txt_nombre_almacen.SetFocus
   End If
End Sub

Private Sub txt_clave_almacen_LostFocus()
   If Trim(txt_clave_almacen) <> "" Then
      rs.Open "select * from tb_almacenes where vcha_alm_almacen_id = '" + txt_clave_almacen + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         txt_nombre_almacen = IIf(IsNull(rs!vcha_alm_nombre), "", rs!vcha_alm_nombre)
      Else
         MsgBox "Clave de almacen incorrecto"
         txt_nombre_almacen = ""
      End If
      rs.Close
   Else
      txt_nombre_almacen = ""
   End If
End Sub

Private Sub txt_fin_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar una fecha"
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

Private Sub txt_fin_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_fin_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_inicio_GotFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = "Presione F5 para seleccionar una fecha"
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

Private Sub txt_inicio_KeyPress(KeyAscii As Integer)
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_inicio_LostFocus()
   Frmmenu2.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub txt_nombre_almacen_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   Else
      Call pro_enfoque(KeyAscii)
   End If
End Sub
