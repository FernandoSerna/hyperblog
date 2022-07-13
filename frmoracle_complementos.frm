VERSION 5.00
Begin VB.Form frmoracle_complementos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contenido del grupo"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   105
      TabIndex        =   45
      Top             =   450
      Width           =   9525
      Begin VB.TextBox txt_descripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2175
         TabIndex        =   6
         Top             =   135
         Width           =   7245
      End
      Begin VB.TextBox txt_codigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1005
         TabIndex        =   5
         Top             =   135
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         TabIndex        =   46
         Top             =   210
         Width           =   825
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   75
      Picture         =   "frmoracle_complementos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      Picture         =   "frmoracle_complementos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      Picture         =   "frmoracle_complementos.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_deshacer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   735
      Picture         =   "frmoracle_complementos.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Deshacer Alt + D"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_carga_mmasiva 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1065
      Picture         =   "frmoracle_complementos.frx":0910
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Carga masiva"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   75
      TabIndex        =   44
      Top             =   345
      Width           =   9510
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   90
      TabIndex        =   39
      Top             =   1080
      Width           =   9540
      Begin VB.TextBox txt_precio_8 
         Height          =   390
         Left            =   3660
         TabIndex        =   22
         Top             =   3690
         Width           =   1065
      End
      Begin VB.TextBox txt_precio_7 
         Height          =   390
         Left            =   3660
         TabIndex        =   20
         Top             =   3255
         Width           =   1065
      End
      Begin VB.TextBox txt_precio_6 
         Height          =   390
         Left            =   3660
         TabIndex        =   18
         Top             =   2835
         Width           =   1065
      End
      Begin VB.TextBox txt_precio_5 
         Height          =   390
         Left            =   3660
         TabIndex        =   16
         Top             =   2400
         Width           =   1065
      End
      Begin VB.TextBox txt_precio_4 
         Height          =   390
         Left            =   3660
         TabIndex        =   14
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox txt_precio_3 
         Height          =   390
         Left            =   3660
         TabIndex        =   12
         Top             =   1545
         Width           =   1065
      End
      Begin VB.TextBox txt_precio_2 
         Height          =   390
         Left            =   3660
         TabIndex        =   10
         Top             =   1125
         Width           =   1065
      End
      Begin VB.TextBox txt_precio_1 
         Height          =   390
         Left            =   3660
         TabIndex        =   8
         Top             =   690
         Width           =   1065
      End
      Begin VB.TextBox txt_precio_16 
         Height          =   390
         Left            =   8355
         TabIndex        =   38
         Top             =   3690
         Width           =   1065
      End
      Begin VB.TextBox txt_complemento_16 
         Height          =   390
         Left            =   4770
         TabIndex        =   37
         Top             =   3690
         Width           =   3570
      End
      Begin VB.TextBox txt_complemento_8 
         Height          =   390
         Left            =   75
         TabIndex        =   21
         Top             =   3690
         Width           =   3570
      End
      Begin VB.TextBox txt_precio_15 
         Height          =   390
         Left            =   8355
         TabIndex        =   36
         Top             =   3255
         Width           =   1065
      End
      Begin VB.TextBox txt_complemento_15 
         Height          =   390
         Left            =   4770
         TabIndex        =   35
         Top             =   3255
         Width           =   3570
      End
      Begin VB.TextBox txt_complemento_7 
         Height          =   390
         Left            =   75
         TabIndex        =   19
         Top             =   3255
         Width           =   3570
      End
      Begin VB.TextBox txt_precio_14 
         Height          =   390
         Left            =   8355
         TabIndex        =   34
         Top             =   2835
         Width           =   1065
      End
      Begin VB.TextBox txt_complemento_14 
         Height          =   390
         Left            =   4770
         TabIndex        =   33
         Top             =   2835
         Width           =   3570
      End
      Begin VB.TextBox txt_complemento_6 
         Height          =   390
         Left            =   75
         TabIndex        =   17
         Top             =   2835
         Width           =   3570
      End
      Begin VB.TextBox txt_precio_13 
         Height          =   390
         Left            =   8355
         TabIndex        =   32
         Top             =   2400
         Width           =   1065
      End
      Begin VB.TextBox txt_complemento_13 
         Height          =   390
         Left            =   4770
         TabIndex        =   31
         Top             =   2400
         Width           =   3570
      End
      Begin VB.TextBox txt_complemento_5 
         Height          =   390
         Left            =   75
         TabIndex        =   15
         Top             =   2400
         Width           =   3570
      End
      Begin VB.TextBox txt_precio_12 
         Height          =   390
         Left            =   8355
         TabIndex        =   30
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox txt_complemento_12 
         Height          =   390
         Left            =   4770
         TabIndex        =   29
         Top             =   1980
         Width           =   3570
      End
      Begin VB.TextBox txt_complemento_4 
         Height          =   390
         Left            =   75
         TabIndex        =   13
         Top             =   1980
         Width           =   3570
      End
      Begin VB.TextBox txt_precio_11 
         Height          =   390
         Left            =   8355
         TabIndex        =   28
         Top             =   1545
         Width           =   1065
      End
      Begin VB.TextBox txt_complemento_11 
         Height          =   390
         Left            =   4770
         TabIndex        =   27
         Top             =   1545
         Width           =   3570
      End
      Begin VB.TextBox txt_complemento_3 
         Height          =   390
         Left            =   75
         TabIndex        =   11
         Top             =   1545
         Width           =   3570
      End
      Begin VB.TextBox txt_precio_10 
         Height          =   390
         Left            =   8355
         TabIndex        =   26
         Top             =   1125
         Width           =   1065
      End
      Begin VB.TextBox txt_complemento_10 
         Height          =   390
         Left            =   4770
         TabIndex        =   25
         Top             =   1125
         Width           =   3570
      End
      Begin VB.TextBox txt_complemento_2 
         Height          =   390
         Left            =   75
         TabIndex        =   9
         Top             =   1125
         Width           =   3570
      End
      Begin VB.TextBox txt_precio_9 
         Height          =   390
         Left            =   8355
         TabIndex        =   24
         Top             =   690
         Width           =   1065
      End
      Begin VB.TextBox txt_complemento_9 
         Height          =   390
         Left            =   4770
         TabIndex        =   23
         Top             =   690
         Width           =   3570
      End
      Begin VB.TextBox txt_complemento_1 
         Height          =   390
         Left            =   75
         TabIndex        =   7
         Top             =   690
         Width           =   3570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   4035
         TabIndex        =   43
         Top             =   345
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   8505
         TabIndex        =   42
         Top             =   345
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Complemento"
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
         Left            =   5670
         TabIndex        =   41
         Top             =   345
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Complemento"
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
         Left            =   930
         TabIndex        =   40
         Top             =   345
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmoracle_complementos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub cmd_carga_mmasiva_Click()
On Error GoTo salir:
   var_codigo = Me.txt_codigo
   strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=C:\REPORTESSID\COMPLEMENTOS_2.XLS"
   If rs.State = 1 Then
      rs.Close
   End If
   rs.Open "select * FROM [COMPLEMENTOS$]", strConnectionString
   While Not rs.EOF
         Me.txt_codigo = rs!CODIGO
         If Me.txt_codigo = "00056850" Then
            Me.txt_codigo = Me.txt_codigo
            'MsgBox IIf(IsNull(rs!precio_5), "", rs!precio_5)
         End If
         Me.txt_complemento_1 = IIf(IsNull(rs!complemento_1), "", rs!complemento_1)
         Me.txt_precio_1 = IIf(IsNull(rs!precio_1), 0, rs!precio_1)
         Me.txt_complemento_2 = IIf(IsNull(rs!complemento_2), "", rs!complemento_2)
         Me.txt_precio_2 = IIf(IsNull(rs!precio_2), 0, rs!precio_2)
         Me.txt_complemento_3 = IIf(IsNull(rs!complemento_3), "", rs!complemento_3)
         Me.txt_precio_3 = IIf(IsNull(rs!precio_3), 0, rs!precio_3)
         Me.txt_complemento_4 = IIf(IsNull(rs!complemento_4), "", rs!complemento_4)
         Me.txt_precio_4 = IIf(IsNull(rs!precio_4), 0, rs!precio_4)
         Me.txt_complemento_5 = IIf(IsNull(rs!complemento_5), "", rs!complemento_5)
         Me.txt_precio_5 = IIf(IsNull(rs!precio_5), 0, rs!precio_5)
         Me.txt_complemento_6 = IIf(IsNull(rs!complemento_6), "", rs!complemento_6)
         Me.txt_precio_6 = IIf(IsNull(rs!precio_6), 0, rs!precio_6)
         Me.txt_complemento_7 = IIf(IsNull(rs!complemento_7), "", rs!complemento_7)
         Me.txt_precio_7 = IIf(IsNull(rs!precio_7), 0, rs!precio_7)
         Me.txt_complemento_8 = IIf(IsNull(rs!complemento_8), "", rs!complemento_8)
         Me.txt_precio_8 = IIf(IsNull(rs!precio_8), 0, rs!precio_8)
         Me.txt_complemento_9 = IIf(IsNull(rs!complemento_9), "", rs!complemento_9)
         Me.txt_precio_9 = IIf(IsNull(rs!precio_9), 0, rs!precio_9)
         Me.txt_complemento_10 = IIf(IsNull(rs!complemento_10), "", rs!complemento_10)
         Me.txt_precio_10 = IIf(IsNull(rs!precio_10), 0, rs!precio_10)
         Me.txt_complemento_11 = IIf(IsNull(rs!complemento_11), "", rs!complemento_11)
         Me.txt_precio_11 = IIf(IsNull(rs!precio_11), 0, rs!precio_11)
         Me.txt_complemento_12 = IIf(IsNull(rs!complemento_12), "", rs!complemento_12)
         Me.txt_precio_12 = IIf(IsNull(rs!precio_12), 0, rs!precio_12)
         Me.txt_complemento_13 = IIf(IsNull(rs!complemento_13), "", rs!complemento_13)
         Me.txt_precio_13 = IIf(IsNull(rs!precio_13), 0, rs!precio_13)
         Me.txt_complemento_14 = IIf(IsNull(rs!complemento_14), "", rs!complemento_14)
         Me.txt_precio_14 = IIf(IsNull(rs!precio_14), 0, rs!precio_14)
         Me.txt_complemento_15 = IIf(IsNull(rs!complemento_15), "", rs!complemento_15)
         Me.txt_precio_15 = IIf(IsNull(rs!precio_15), 0, rs!precio_15)
         Me.txt_complemento_16 = IIf(IsNull(rs!complemento_16), "", rs!complemento_16)
         Me.txt_precio_16 = IIf(IsNull(rs!precio_16), 0, rs!precio_16)
         If Not IsNumeric(Me.txt_precio_1) Then
            Me.txt_precio_1 = 0
         End If
         If Not IsNumeric(Me.txt_precio_2) Then
            Me.txt_precio_2 = 0
         End If
         If Not IsNumeric(Me.txt_precio_3) Then
            Me.txt_precio_3 = 0
         End If
         If Not IsNumeric(Me.txt_precio_4) Then
            Me.txt_precio_4 = 0
         End If
         If Not IsNumeric(Me.txt_precio_5) Then
            Me.txt_precio_5 = 0
         End If
         If Not IsNumeric(Me.txt_precio_6) Then
            Me.txt_precio_6 = 0
         End If
         If Not IsNumeric(Me.txt_precio_7) Then
            Me.txt_precio_7 = 0
         End If
         If Not IsNumeric(Me.txt_precio_8) Then
            Me.txt_precio_8 = 0
         End If
         If Not IsNumeric(Me.txt_precio_9) Then
            Me.txt_precio_9 = 0
         End If
         If Not IsNumeric(Me.txt_precio_10) Then
            Me.txt_precio_10 = 0
         End If
         If Not IsNumeric(Me.txt_precio_11) Then
            Me.txt_precio_11 = 0
         End If
         If Not IsNumeric(Me.txt_precio_12) Then
            Me.txt_precio_12 = 0
         End If
         If Not IsNumeric(Me.txt_precio_13) Then
            Me.txt_precio_13 = 0
         End If
         If Not IsNumeric(Me.txt_precio_14) Then
            Me.txt_precio_14 = 0
         End If
         If Not IsNumeric(Me.txt_precio_15) Then
            Me.txt_precio_15 = 0
         End If
         If Not IsNumeric(Me.txt_precio_16) Then
            Me.txt_precio_16 = 0
         End If
         strconsulta = "UPDATE xxvia_tb_complementos_pk_list SET COMPLEMENTO_1 = ?, PRECIO_1 = ?, COMPLEMENTO_2 = ?, PRECIO_2 = ?, COMPLEMENTO_3 = ?, PRECIO_3 = ?, COMPLEMENTO_4 = ?, PRECIO_4 = ?, COMPLEMENTO_5 = ?, PRECIO_5 = ?, COMPLEMENTO_6 = ?, PRECIO_6 = ?, COMPLEMENTO_7 = ?, PRECIO_7 = ?, COMPLEMENTO_8 = ?, PRECIO_8 = ?, COMPLEMENTO_9 = ?, PRECIO_9 = ?, COMPLEMENTO_10 = ?, PRECIO_10 = ?, COMPLEMENTO_11 = ?, PRECIO_11 = ?, COMPLEMENTO_12 = ?, PRECIO_12 = ?, COMPLEMENTO_13 = ?, PRECIO_13 = ?, COMPLEMENTO_14 = ?, PRECIO_14 = ?, COMPLEMENTO_15 = ?, PRECIO_15 = ?, COMPLEMENTO_16 = ?, PRECIO_16 = ? where codigo = ?"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = strconsulta
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_1)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_1))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_2)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_2))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_3)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_3))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_4)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_4))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_5)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_5))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_6)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_6))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_7)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_7))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_8)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_8))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_9)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_9))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_10)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_10))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_11)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_11))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_12)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_12))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_13)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_13))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_14)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_14))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_15)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_15))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_16)
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_16))
              .Parameters.Append parametro
              Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
              .Parameters.Append parametro
         End With
         Set rsaux = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         rs.MoveNext
   Wend
   rs.Close
   
   Me.txt_codigo = var_codigo
   strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
   With comandoORA
        .ActiveConnection = cnnoracle_4
        .CommandType = adCmdText
        .CommandText = strconsulta
        Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
        .Parameters.Append parametro
   End With
   Set rs = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   If Not rs.EOF Then
      Me.txt_complemento_1 = IIf(IsNull(rs!complemento_1), "", rs!complemento_1)
      Me.txt_precio_1 = IIf(IsNull(rs!precio_1), 0, rs!precio_1)
      Me.txt_complemento_2 = IIf(IsNull(rs!complemento_2), "", rs!complemento_2)
      Me.txt_precio_2 = IIf(IsNull(rs!precio_2), 0, rs!precio_2)
      Me.txt_complemento_3 = IIf(IsNull(rs!complemento_3), "", rs!complemento_3)
      Me.txt_precio_3 = IIf(IsNull(rs!precio_3), 0, rs!precio_3)
      Me.txt_complemento_4 = IIf(IsNull(rs!complemento_4), "", rs!complemento_4)
      Me.txt_precio_4 = IIf(IsNull(rs!precio_4), 0, rs!precio_4)
      Me.txt_complemento_5 = IIf(IsNull(rs!complemento_5), "", rs!complemento_5)
      Me.txt_precio_5 = IIf(IsNull(rs!precio_5), 0, rs!precio_5)
      Me.txt_complemento_6 = IIf(IsNull(rs!complemento_6), "", rs!complemento_6)
      Me.txt_precio_6 = IIf(IsNull(rs!precio_6), 0, rs!precio_6)
      Me.txt_complemento_7 = IIf(IsNull(rs!complemento_7), "", rs!complemento_7)
      Me.txt_precio_7 = IIf(IsNull(rs!precio_7), 0, rs!precio_7)
      Me.txt_complemento_8 = IIf(IsNull(rs!complemento_8), "", rs!complemento_8)
      Me.txt_precio_8 = IIf(IsNull(rs!precio_8), 0, rs!precio_8)
      Me.txt_complemento_9 = IIf(IsNull(rs!complemento_9), "", rs!complemento_9)
      Me.txt_precio_9 = IIf(IsNull(rs!precio_9), 0, rs!precio_9)
      Me.txt_complemento_10 = IIf(IsNull(rs!complemento_10), "", rs!complemento_10)
      Me.txt_precio_10 = IIf(IsNull(rs!precio_10), 0, rs!precio_10)
      Me.txt_complemento_11 = IIf(IsNull(rs!complemento_11), "", rs!complemento_11)
      Me.txt_precio_11 = IIf(IsNull(rs!precio_11), 0, rs!precio_11)
      Me.txt_complemento_12 = IIf(IsNull(rs!complemento_12), "", rs!complemento_12)
      Me.txt_precio_12 = IIf(IsNull(rs!precio_12), 0, rs!precio_12)
      Me.txt_complemento_13 = IIf(IsNull(rs!complemento_13), "", rs!complemento_13)
      Me.txt_precio_13 = IIf(IsNull(rs!precio_13), 0, rs!precio_13)
      Me.txt_complemento_14 = IIf(IsNull(rs!complemento_14), "", rs!complemento_14)
      Me.txt_precio_14 = IIf(IsNull(rs!precio_14), 0, rs!precio_14)
      Me.txt_complemento_15 = IIf(IsNull(rs!complemento_15), "", rs!complemento_15)
      Me.txt_precio_15 = IIf(IsNull(rs!precio_15), 0, rs!precio_15)
      Me.txt_complemento_16 = IIf(IsNull(rs!complemento_16), "", rs!complemento_16)
      Me.txt_precio_16 = IIf(IsNull(rs!precio_16), 0, rs!precio_16)
   End If
   rs.Close
   
   
   MsgBox "Se a terminado el proceso de carga masiva", vbOKOnly, "ATENCION"
   Exit Sub
salir:
   MsgBox "Error al cargar el archivo, verifique que el archivo se llame complementos_2, la hoja se llame complementos, que los nombres de las columnas sean codigo, complemento_1, precio_1... complemento_16, precio_16 y que el archivo se guarde en c:\reportessid\", vbOKOnly, "ATENCION"

End Sub

Private Sub cmd_guardar_Click()
   var_si = MsgBox("¿Desea guardar los complementos?", vbYesNo, "ATENCION")
   If var_si = 6 Then
      If Not IsNumeric(Me.txt_precio_1) Then
         Me.txt_precio_1 = 0
      End If
      If Not IsNumeric(Me.txt_precio_2) Then
         Me.txt_precio_2 = 0
      End If
      If Not IsNumeric(Me.txt_precio_3) Then
         Me.txt_precio_3 = 0
      End If
      If Not IsNumeric(Me.txt_precio_4) Then
         Me.txt_precio_4 = 0
      End If
      If Not IsNumeric(Me.txt_precio_5) Then
         Me.txt_precio_5 = 0
      End If
      If Not IsNumeric(Me.txt_precio_6) Then
         Me.txt_precio_6 = 0
      End If
      If Not IsNumeric(Me.txt_precio_7) Then
         Me.txt_precio_7 = 0
      End If
      If Not IsNumeric(Me.txt_precio_8) Then
         Me.txt_precio_8 = 0
      End If
      If Not IsNumeric(Me.txt_precio_9) Then
         Me.txt_precio_9 = 0
      End If
      If Not IsNumeric(Me.txt_precio_10) Then
         Me.txt_precio_10 = 0
      End If
      If Not IsNumeric(Me.txt_precio_11) Then
         Me.txt_precio_11 = 0
      End If
      If Not IsNumeric(Me.txt_precio_12) Then
         Me.txt_precio_12 = 0
      End If
      If Not IsNumeric(Me.txt_precio_13) Then
         Me.txt_precio_13 = 0
      End If
      If Not IsNumeric(Me.txt_precio_14) Then
         Me.txt_precio_14 = 0
      End If
      If Not IsNumeric(Me.txt_precio_15) Then
         Me.txt_precio_15 = 0
      End If
      If Not IsNumeric(Me.txt_precio_16) Then
         Me.txt_precio_16 = 0
      End If
      strconsulta = "UPDATE xxvia_tb_complementos_pk_list SET COMPLEMENTO_1 = ?, PRECIO_1 = ?, COMPLEMENTO_2 = ?, PRECIO_2 = ?, COMPLEMENTO_3 = ?, PRECIO_3 = ?, COMPLEMENTO_4 = ?, PRECIO_4 = ?, COMPLEMENTO_5 = ?, PRECIO_5 = ?, COMPLEMENTO_6 = ?, PRECIO_6 = ?, COMPLEMENTO_7 = ?, PRECIO_7 = ?, COMPLEMENTO_8 = ?, PRECIO_8 = ?, COMPLEMENTO_9 = ?, PRECIO_9 = ?, COMPLEMENTO_10 = ?, PRECIO_10 = ?, COMPLEMENTO_11 = ?, PRECIO_11 = ?, COMPLEMENTO_12 = ?, PRECIO_12 = ?, COMPLEMENTO_13 = ?, PRECIO_13 = ?, COMPLEMENTO_14 = ?, PRECIO_14 = ?, COMPLEMENTO_15 = ?, PRECIO_15 = ?, COMPLEMENTO_16 = ?, PRECIO_16 = ? where codigo = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = strconsulta
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_1)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_1))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_2)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_2))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_3)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_3))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_4)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_4))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_5)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_5))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_6)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_6))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_7)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_7))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_8)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_8))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_9)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_9))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_10)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_10))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_11)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_11))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_12)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_12))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_13)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_13))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_14)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_14))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_15)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_15))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_complemento_16)
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 200, CDbl(Me.txt_precio_16))
           .Parameters.Append parametro
           Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
           .Parameters.Append parametro
      End With
      Set rs = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      
   End If
End Sub

Private Sub cmd_nuevo_Click()
   Me.txt_complemento_1 = ""
   Me.txt_complemento_2 = ""
   Me.txt_complemento_3 = ""
   Me.txt_complemento_4 = ""
   Me.txt_complemento_5 = ""
   Me.txt_complemento_6 = ""
   Me.txt_complemento_7 = ""
   Me.txt_complemento_8 = ""
   Me.txt_complemento_9 = ""
   Me.txt_complemento_10 = ""
   Me.txt_complemento_11 = ""
   Me.txt_complemento_12 = ""
   Me.txt_complemento_13 = ""
   Me.txt_complemento_14 = ""
   Me.txt_complemento_15 = ""
   Me.txt_complemento_16 = ""
   Me.txt_precio_1 = 0
   Me.txt_precio_2 = 0
   Me.txt_precio_3 = 0
   Me.txt_precio_4 = 0
   Me.txt_precio_5 = 0
   Me.txt_precio_6 = 0
   Me.txt_precio_7 = 0
   Me.txt_precio_8 = 0
   Me.txt_precio_9 = 0
   Me.txt_precio_10 = 0
   Me.txt_precio_11 = 0
   Me.txt_precio_12 = 0
   Me.txt_precio_13 = 0
   Me.txt_precio_14 = 0
   Me.txt_precio_15 = 0
   Me.txt_precio_16 = 0
   Me.txt_complemento_1.SetFocus
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.txt_codigo = var_codigo_complemento
   Me.txt_descripcion = var_descripcion_complemento
   strconsulta = "select * from xxvia_tb_complementos_pk_list where codigo = ?"
   With comandoORA
        .ActiveConnection = cnnoracle_4
        .CommandType = adCmdText
        .CommandText = strconsulta
        Set parametro = .CreateParameter(, adVarChar, adParamInput, 200, Me.txt_codigo)
        .Parameters.Append parametro
   End With
   Set rs = comandoORA.execute
   Set comandoORA = Nothing
   Set parametro = Nothing
   If Not rs.EOF Then
      Me.txt_complemento_1 = IIf(IsNull(rs!complemento_1), "", rs!complemento_1)
      Me.txt_precio_1 = IIf(IsNull(rs!precio_1), 0, rs!precio_1)
      Me.txt_complemento_2 = IIf(IsNull(rs!complemento_2), "", rs!complemento_2)
      Me.txt_precio_2 = IIf(IsNull(rs!precio_2), 0, rs!precio_2)
      Me.txt_complemento_3 = IIf(IsNull(rs!complemento_3), "", rs!complemento_3)
      Me.txt_precio_3 = IIf(IsNull(rs!precio_3), 0, rs!precio_3)
      Me.txt_complemento_4 = IIf(IsNull(rs!complemento_4), "", rs!complemento_4)
      Me.txt_precio_4 = IIf(IsNull(rs!precio_4), 0, rs!precio_4)
      Me.txt_complemento_5 = IIf(IsNull(rs!complemento_5), "", rs!complemento_5)
      Me.txt_precio_5 = IIf(IsNull(rs!precio_5), 0, rs!precio_5)
      Me.txt_complemento_6 = IIf(IsNull(rs!complemento_6), "", rs!complemento_6)
      Me.txt_precio_6 = IIf(IsNull(rs!precio_6), 0, rs!precio_6)
      Me.txt_complemento_7 = IIf(IsNull(rs!complemento_7), "", rs!complemento_7)
      Me.txt_precio_7 = IIf(IsNull(rs!precio_7), 0, rs!precio_7)
      Me.txt_complemento_8 = IIf(IsNull(rs!complemento_8), "", rs!complemento_8)
      Me.txt_precio_8 = IIf(IsNull(rs!precio_8), 0, rs!precio_8)
      Me.txt_complemento_9 = IIf(IsNull(rs!complemento_9), "", rs!complemento_9)
      Me.txt_precio_9 = IIf(IsNull(rs!precio_9), 0, rs!precio_9)
      Me.txt_complemento_10 = IIf(IsNull(rs!complemento_10), "", rs!complemento_10)
      Me.txt_precio_10 = IIf(IsNull(rs!precio_10), 0, rs!precio_10)
      Me.txt_complemento_11 = IIf(IsNull(rs!complemento_11), "", rs!complemento_11)
      Me.txt_precio_11 = IIf(IsNull(rs!precio_11), 0, rs!precio_11)
      Me.txt_complemento_12 = IIf(IsNull(rs!complemento_12), "", rs!complemento_12)
      Me.txt_precio_12 = IIf(IsNull(rs!precio_12), 0, rs!precio_12)
      Me.txt_complemento_13 = IIf(IsNull(rs!complemento_13), "", rs!complemento_13)
      Me.txt_precio_13 = IIf(IsNull(rs!precio_13), 0, rs!precio_13)
      Me.txt_complemento_14 = IIf(IsNull(rs!complemento_14), "", rs!complemento_14)
      Me.txt_precio_14 = IIf(IsNull(rs!precio_14), 0, rs!precio_14)
      Me.txt_complemento_15 = IIf(IsNull(rs!complemento_15), "", rs!complemento_15)
      Me.txt_precio_15 = IIf(IsNull(rs!precio_15), 0, rs!precio_15)
      Me.txt_complemento_16 = IIf(IsNull(rs!complemento_16), "", rs!complemento_16)
      Me.txt_precio_16 = IIf(IsNull(rs!precio_16), 0, rs!precio_16)
   End If
   rs.Close
   
End Sub

Private Sub txt_complemento_1_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_10_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_11_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_12_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_13_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_14_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_15_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_16_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_2_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_3_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_4_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_5_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_6_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_7_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_8_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_complemento_9_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_1_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_10_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_11_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_12_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_13_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_14_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_15_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.cmd_guardar.SetFocus
    End If
End Sub

Private Sub txt_precio_2_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_3_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_4_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_5_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_6_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_7_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_8_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub

Private Sub txt_precio_9_KeyPress(KeyAscii As Integer)
    Call pro_enfoque(KeyAscii)
End Sub
