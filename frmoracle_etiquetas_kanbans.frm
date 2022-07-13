VERSION 5.00
Begin VB.Form frmoracle_etiquetas_kanbans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de etiquetas Kanban"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   11220
      Picture         =   "frmoracle_etiquetas_kanbans.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_etiquetas_kanbans.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Nuevo Movimiento"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmoracle_etiquetas_kanbans.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Imprimir"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   6870
      Left            =   30
      TabIndex        =   9
      Top             =   375
      Width           =   11535
      Begin VB.TextBox txt_ubicacion_5_4 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   71
         Top             =   6420
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_4_4 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   70
         Top             =   6105
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_3_4 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   69
         Top             =   5775
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_2_4 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   68
         Top             =   5445
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_1_4 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   67
         Top             =   5115
         Width           =   1905
      End
      Begin VB.TextBox txt_cantidad_4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   6750
         TabIndex        =   7
         Top             =   4650
         Width           =   1485
      End
      Begin VB.TextBox txt_descripcion_4 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   66
         Top             =   4320
         Width           =   4710
      End
      Begin VB.TextBox txt_codigo_4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6750
         TabIndex        =   6
         Top             =   3825
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_5_3 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   56
         Top             =   6420
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_4_3 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   55
         Top             =   6090
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_3_3 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   54
         Top             =   5760
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_2_3 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   53
         Top             =   5430
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_1_3 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   52
         Top             =   5115
         Width           =   1905
      End
      Begin VB.TextBox txt_cantidad_3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1020
         TabIndex        =   5
         Top             =   4650
         Width           =   1485
      End
      Begin VB.TextBox txt_descripcion_3 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   51
         Top             =   4320
         Width           =   4605
      End
      Begin VB.TextBox txt_codigo_3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1020
         TabIndex        =   4
         Top             =   3825
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_5_2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   41
         Top             =   3075
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_4_2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   40
         Top             =   2745
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_3_2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   39
         Top             =   2415
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_2_2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   38
         Top             =   2085
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_1_2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   37
         Top             =   1755
         Width           =   1905
      End
      Begin VB.TextBox txt_cantidad_2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   6750
         TabIndex        =   3
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox txt_descripcion_2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6750
         TabIndex        =   36
         Top             =   960
         Width           =   4710
      End
      Begin VB.TextBox txt_codigo_2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6750
         TabIndex        =   2
         Top             =   465
         Width           =   1935
      End
      Begin VB.TextBox txt_ubicacion_5_1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   26
         Top             =   3075
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_4_1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   25
         Top             =   2745
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_3_1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   24
         Top             =   2415
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_2_1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   23
         Top             =   2085
         Width           =   1905
      End
      Begin VB.TextBox txt_ubicacion_1_1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   22
         Top             =   1755
         Width           =   1905
      End
      Begin VB.TextBox txt_cantidad_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1020
         TabIndex        =   1
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox txt_descripcion_1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1020
         TabIndex        =   21
         Top             =   960
         Width           =   4605
      End
      Begin VB.TextBox txt_codigo_1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1020
         TabIndex        =   0
         Top             =   465
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   6855
         Left            =   5700
         TabIndex        =   11
         Top             =   0
         Width           =   30
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   15
         TabIndex        =   10
         Top             =   3330
         Width           =   11490
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 5:"
         Height          =   195
         Left            =   5805
         TabIndex        =   65
         Top             =   6480
         Width           =   900
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 4:"
         Height          =   195
         Left            =   5805
         TabIndex        =   64
         Top             =   6150
         Width           =   900
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 3:"
         Height          =   195
         Left            =   5805
         TabIndex        =   63
         Top             =   5820
         Width           =   900
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 2:"
         Height          =   195
         Left            =   5805
         TabIndex        =   62
         Top             =   5490
         Width           =   900
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 1:"
         Height          =   195
         Left            =   5805
         TabIndex        =   61
         Top             =   5175
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5805
         TabIndex        =   60
         Top             =   4770
         Width           =   675
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   5805
         TabIndex        =   59
         Top             =   4380
         Width           =   885
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   5805
         TabIndex        =   58
         Top             =   3960
         Width           =   540
      End
      Begin VB.Label Label28 
         BackColor       =   &H000000FF&
         Caption         =   " Etiqueta 4"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5730
         TabIndex        =   57
         Top             =   3465
         Width           =   5760
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 5:"
         Height          =   195
         Left            =   75
         TabIndex        =   50
         Top             =   6480
         Width           =   900
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 4:"
         Height          =   195
         Left            =   75
         TabIndex        =   49
         Top             =   6150
         Width           =   900
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 3:"
         Height          =   195
         Left            =   75
         TabIndex        =   48
         Top             =   5820
         Width           =   900
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 2:"
         Height          =   195
         Left            =   75
         TabIndex        =   47
         Top             =   5490
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 1:"
         Height          =   195
         Left            =   75
         TabIndex        =   46
         Top             =   5175
         Width           =   900
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   75
         TabIndex        =   45
         Top             =   4770
         Width           =   675
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   75
         TabIndex        =   44
         Top             =   4380
         Width           =   885
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   75
         TabIndex        =   43
         Top             =   3960
         Width           =   540
      End
      Begin VB.Label Label19 
         BackColor       =   &H000000FF&
         Caption         =   " Etiqueta 3"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   75
         TabIndex        =   42
         Top             =   3465
         Width           =   5640
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 5:"
         Height          =   195
         Left            =   5805
         TabIndex        =   35
         Top             =   3135
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 4:"
         Height          =   195
         Left            =   5805
         TabIndex        =   34
         Top             =   2805
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 3:"
         Height          =   195
         Left            =   5805
         TabIndex        =   33
         Top             =   2475
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 2:"
         Height          =   195
         Left            =   5805
         TabIndex        =   32
         Top             =   2145
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 1:"
         Height          =   195
         Left            =   5805
         TabIndex        =   31
         Top             =   1815
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5805
         TabIndex        =   30
         Top             =   1410
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   5805
         TabIndex        =   29
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   5805
         TabIndex        =   28
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000FF&
         Caption         =   " Etiqueta 2"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5745
         TabIndex        =   27
         Top             =   135
         Width           =   5745
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 5:"
         Height          =   195
         Left            =   75
         TabIndex        =   20
         Top             =   3135
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 4:"
         Height          =   195
         Left            =   75
         TabIndex        =   19
         Top             =   2805
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 3:"
         Height          =   195
         Left            =   75
         TabIndex        =   18
         Top             =   2475
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 2:"
         Height          =   195
         Left            =   75
         TabIndex        =   17
         Top             =   2145
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion 1:"
         Height          =   195
         Left            =   75
         TabIndex        =   16
         Top             =   1815
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   75
         TabIndex        =   15
         Top             =   1410
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   75
         TabIndex        =   14
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   75
         TabIndex        =   13
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Caption         =   " Etiqueta 1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   12
         Top             =   135
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   15
      TabIndex        =   8
      Top             =   345
      Width           =   11535
   End
End
Attribute VB_Name = "frmoracle_etiquetas_kanbans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
   cnn.BeginTrans
   rs.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_ETIQUETAS_KANBAN", cnn, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
      var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
   Else
      var_consecutivo = 1
   End If
   rs.Close
   var_no = ""
   If Me.txt_codigo_1 <> "" Then
      If Not IsNumeric(Me.txt_cantidad_1) Then
         var_no = "1"
      End If
   End If
   If Me.txt_codigo_2 <> "" Then
      If Not IsNumeric(Me.txt_cantidad_2) Then
         If var_no = "" Then
            var_no = "2"
         Else
            var_no = var_no + ", 2"
         End If
      End If
   End If
   If Me.txt_codigo_3 <> "" Then
      If Not IsNumeric(Me.txt_cantidad_3) Then
         If var_no = "" Then
            var_no = "3"
         Else
            var_no = var_no + ", 3"
         End If
      End If
   End If
   If Me.txt_codigo_4 <> "" Then
      If Not IsNumeric(Me.txt_cantidad_4) Then
         If var_no = "" Then
            var_no = "4"
         Else
            var_no = var_no + ", 4"
         End If
      End If
   End If
   If var_no = "" Then
      If Me.txt_cantidad_1 = "" Then
         Me.txt_cantidad_1 = 0
      End If
      If Me.txt_cantidad_2 = "" Then
         Me.txt_cantidad_2 = 0
      End If
      If Me.txt_cantidad_3 = "" Then
         Me.txt_cantidad_3 = 0
      End If
      If Me.txt_cantidad_4 = "" Then
         Me.txt_cantidad_4 = 0
      End If
      var_cadena = "INSERT INTO TB_TEMP_ORACLE_ETIQUETAS_KANBAN (INTE_TEM_CONSECUTIVO, CODIGO_1, DESCRIPCION_1, CANTIDAD_1, UBICACION_1_1, UBICACION_2_1, UBICACION_3_1, UBICACION_4_1, UBICACION_5_1, CODIGO_2, DESCRIPCION_2, CANTIDAD_2, UBICACION_1_2, UBICACION_2_2, UBICACION_3_2, UBICACION_4_2, UBICACION_5_2, CODIGO_3, DESCRIPCION_3, CANTIDAD_3, UBICACION_1_3, UBICACION_2_3, UBICACION_3_3, UBICACION_4_3, UBICACION_5_3, CODIGO_4, DESCRIPCION_4, CANTIDAD_4, UBICACION_1_4, UBICACION_2_4, UBICACION_3_4, UBICACION_4_4, UBICACION_5_4) VALUES (" + CStr(var_consecutivo)
      var_cadena = var_cadena + ",'" + Me.txt_codigo_1 + "','" + Me.txt_descripcion_1 + "'," + Me.txt_cantidad_1 + ",'" + Me.txt_ubicacion_1_1 + "','" + Me.txt_ubicacion_2_1 + "','" + Me.txt_ubicacion_3_1 + "','" + Me.txt_ubicacion_4_1 + "','" + Me.txt_ubicacion_5_1
      var_cadena = var_cadena + "','" + Me.txt_codigo_2 + "','" + Me.txt_descripcion_2 + "'," + Me.txt_cantidad_2 + ",'" + Me.txt_ubicacion_1_2 + "','" + Me.txt_ubicacion_2_2 + "','" + Me.txt_ubicacion_3_2 + "','" + Me.txt_ubicacion_4_2 + "','" + Me.txt_ubicacion_5_2
      var_cadena = var_cadena + "','" + Me.txt_codigo_3 + "','" + Me.txt_descripcion_3 + "'," + Me.txt_cantidad_3 + ",'" + Me.txt_ubicacion_1_3 + "','" + Me.txt_ubicacion_2_3 + "','" + Me.txt_ubicacion_3_3 + "','" + Me.txt_ubicacion_4_3 + "','" + Me.txt_ubicacion_5_3
      var_cadena = var_cadena + "','" + Me.txt_codigo_4 + "','" + Me.txt_descripcion_4 + "'," + Me.txt_cantidad_4 + ",'" + Me.txt_ubicacion_1_4 + "','" + Me.txt_ubicacion_2_4 + "','" + Me.txt_ubicacion_3_4 + "','" + Me.txt_ubicacion_4_4 + "','" + Me.txt_ubicacion_5_4 + "')"
      'MsgBox var_cadena
      rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
      cnn.CommitTrans
      Set reporte = appl.OpenReport(App.Path + "\rep_oracle_etiquetas_kanban.rpt")
      var_cadena = "{TB_TEMP_ORACLE_ETIQUETAS_KANBAN.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
      reporte.RecordSelectionFormula = var_cadena
      frmvistasprevias.cr.ReportSource = reporte
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Etiquetas kanban"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      rs.Open "delete from TB_TEMP_ORACLE_ETIQUETAS_KANBAN where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
   Else
      MsgBox "Cantidades incorrectas en las etiquetas " + var_no, vbOKOnly, "ATENCION"
   End If

End Sub

Private Sub cmd_nuevo_Click()
   Call pro_limpiatextos(Me)
   Me.txt_codigo_1.SetFocus
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Top = 0
   Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_cantidad_1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cantidad_2_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cantidad_3_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_cantidad_4_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Me.txt_codigo_1) = 5 Then
         Me.txt_codigo_1 = "000" + Me.txt_codigo_1
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_1_LostFocus()
   If Me.txt_codigo_1 <> "" Then
      rs.Open "select attribute2 ubicacion1, attribute3  ubicacion2, attribute4  ubicacion3, attribute5  ubicacion4, attribute6 ubicacion5, attribute7 ubicacion6, inventory_item_id item_id, segment1 item_number, description item_description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and segment1 = '" + Me.txt_codigo_1 + "' order by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion_1 = rs!item_description
         Me.txt_ubicacion_1_1 = IIf(IsNull(rs!UBICACION1), "", rs!UBICACION1)
         Me.txt_ubicacion_2_1 = IIf(IsNull(rs!UBICACION2), "", rs!UBICACION2)
         Me.txt_ubicacion_3_1 = IIf(IsNull(rs!UBICACION3), "", rs!UBICACION3)
         Me.txt_ubicacion_4_1 = IIf(IsNull(rs!UBICACION4), "", rs!UBICACION4)
         Me.txt_ubicacion_5_1 = IIf(IsNull(rs!UBICACION5), "", rs!UBICACION5)
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_codigo_1 = ""
         Me.txt_cantidad_1 = ""
         Me.txt_descripcion_1 = ""
         Me.txt_ubicacion_1_1 = ""
         Me.txt_ubicacion_2_1 = ""
         Me.txt_ubicacion_3_1 = ""
         Me.txt_ubicacion_4_1 = ""
         Me.txt_ubicacion_5_1 = ""
      End If
      rs.Close
   Else
      Me.txt_cantidad_1 = ""
      Me.txt_descripcion_1 = ""
      Me.txt_ubicacion_1_1 = ""
      Me.txt_ubicacion_2_1 = ""
      Me.txt_ubicacion_3_1 = ""
      Me.txt_ubicacion_4_1 = ""
      Me.txt_ubicacion_5_1 = ""
   End If
End Sub

Private Sub txt_codigo_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Me.txt_codigo_2) = 5 Then
         Me.txt_codigo_2 = "000" + Me.txt_codigo_2
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_2_LostFocus()
   If Me.txt_codigo_2 <> "" Then
      rs.Open "select attribute2 ubicacion1, attribute3  ubicacion2, attribute4  ubicacion3, attribute5  ubicacion4, attribute6 ubicacion5, attribute7 ubicacion6, inventory_item_id item_id, segment1 item_number, description item_description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and segment1 = '" + Me.txt_codigo_2 + "' order by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion_2 = rs!item_description
         Me.txt_ubicacion_1_2 = IIf(IsNull(rs!UBICACION1), "", rs!UBICACION1)
         Me.txt_ubicacion_2_2 = IIf(IsNull(rs!UBICACION2), "", rs!UBICACION2)
         Me.txt_ubicacion_3_2 = IIf(IsNull(rs!UBICACION3), "", rs!UBICACION3)
         Me.txt_ubicacion_4_2 = IIf(IsNull(rs!UBICACION4), "", rs!UBICACION4)
         Me.txt_ubicacion_5_2 = IIf(IsNull(rs!UBICACION5), "", rs!UBICACION5)
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_codigo_2 = ""
         Me.txt_cantidad_2 = ""
         Me.txt_descripcion_2 = ""
         Me.txt_ubicacion_1_2 = ""
         Me.txt_ubicacion_2_2 = ""
         Me.txt_ubicacion_3_2 = ""
         Me.txt_ubicacion_4_2 = ""
         Me.txt_ubicacion_5_2 = ""
      End If
      rs.Close
   Else
      Me.txt_cantidad_2 = ""
      Me.txt_descripcion_2 = ""
      Me.txt_ubicacion_1_2 = ""
      Me.txt_ubicacion_2_2 = ""
      Me.txt_ubicacion_3_2 = ""
      Me.txt_ubicacion_4_2 = ""
      Me.txt_ubicacion_5_2 = ""
   End If
End Sub

Private Sub txt_codigo_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Me.txt_codigo_3) = 5 Then
         Me.txt_codigo_3 = "000" + Me.txt_codigo_3
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_3_LostFocus()
   If Me.txt_codigo_3 <> "" Then
      rs.Open "select attribute2 ubicacion1, attribute3  ubicacion2, attribute4  ubicacion3, attribute5  ubicacion4, attribute6 ubicacion5, attribute7 ubicacion6, inventory_item_id item_id, segment1 item_number, description item_description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and segment1 = '" + Me.txt_codigo_3 + "' order by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion_3 = rs!item_description
         Me.txt_ubicacion_1_3 = IIf(IsNull(rs!UBICACION1), "", rs!UBICACION1)
         Me.txt_ubicacion_2_3 = IIf(IsNull(rs!UBICACION2), "", rs!UBICACION2)
         Me.txt_ubicacion_3_3 = IIf(IsNull(rs!UBICACION3), "", rs!UBICACION3)
         Me.txt_ubicacion_4_3 = IIf(IsNull(rs!UBICACION4), "", rs!UBICACION4)
         Me.txt_ubicacion_5_3 = IIf(IsNull(rs!UBICACION5), "", rs!UBICACION5)
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_codigo_3 = ""
         Me.txt_cantidad_3 = ""
         Me.txt_descripcion_3 = ""
         Me.txt_ubicacion_1_3 = ""
         Me.txt_ubicacion_2_3 = ""
         Me.txt_ubicacion_3_3 = ""
         Me.txt_ubicacion_4_3 = ""
         Me.txt_ubicacion_5_3 = ""
      End If
      rs.Close
   Else
      Me.txt_cantidad_3 = ""
      Me.txt_descripcion_3 = ""
      Me.txt_ubicacion_1_3 = ""
      Me.txt_ubicacion_2_3 = ""
      Me.txt_ubicacion_3_3 = ""
      Me.txt_ubicacion_4_3 = ""
      Me.txt_ubicacion_5_3 = ""
   End If
End Sub

Private Sub txt_codigo_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Me.txt_codigo_4) = 5 Then
         Me.txt_codigo_4 = "000" + Me.txt_codigo_4
      End If
   End If
   Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_codigo_4_LostFocus()
   If Me.txt_codigo_4 <> "" Then
      rs.Open "select attribute2 ubicacion1, attribute3  ubicacion2, attribute4  ubicacion3, attribute5  ubicacion4, attribute6 ubicacion5, attribute7 ubicacion6, inventory_item_id item_id, segment1 item_number, description item_description from  xxvia_system_items_b where  organization_id  = " + var_unidad_organizacional + " and segment1 = '" + Me.txt_codigo_4 + "' order by segment1", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Me.txt_descripcion_4 = rs!item_description
         Me.txt_ubicacion_1_4 = IIf(IsNull(rs!UBICACION1), "", rs!UBICACION1)
         Me.txt_ubicacion_2_4 = IIf(IsNull(rs!UBICACION2), "", rs!UBICACION2)
         Me.txt_ubicacion_3_4 = IIf(IsNull(rs!UBICACION3), "", rs!UBICACION3)
         Me.txt_ubicacion_4_4 = IIf(IsNull(rs!UBICACION4), "", rs!UBICACION4)
         Me.txt_ubicacion_5_4 = IIf(IsNull(rs!UBICACION5), "", rs!UBICACION5)
      Else
         MsgBox "El artículo no existe", vbOKOnly, "ATENCION"
         Me.txt_codigo_4 = ""
         Me.txt_cantidad_4 = ""
         Me.txt_descripcion_4 = ""
         Me.txt_ubicacion_1_4 = ""
         Me.txt_ubicacion_2_4 = ""
         Me.txt_ubicacion_3_4 = ""
         Me.txt_ubicacion_4_4 = ""
         Me.txt_ubicacion_5_4 = ""
      End If
      rs.Close
   Else
      Me.txt_cantidad_4 = ""
      Me.txt_descripcion_4 = ""
      Me.txt_ubicacion_1_4 = ""
      Me.txt_ubicacion_2_4 = ""
      Me.txt_ubicacion_3_4 = ""
      Me.txt_ubicacion_4_4 = ""
      Me.txt_ubicacion_5_4 = ""
   End If
End Sub
