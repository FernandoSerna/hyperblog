VERSION 5.00
Begin VB.Form frmoracle_ubicaciones_kanbans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubicaciones Kanbans"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3060
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   5595
      Begin VB.TextBox txt_ubicacion_6 
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
         Left            =   1335
         TabIndex        =   12
         Top             =   2505
         Width           =   4110
      End
      Begin VB.TextBox txt_ubicacion_5 
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
         Left            =   1335
         TabIndex        =   10
         Top             =   2055
         Width           =   4110
      End
      Begin VB.TextBox txt_ubicacion_4 
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
         Left            =   1335
         TabIndex        =   8
         Top             =   1605
         Width           =   4110
      End
      Begin VB.TextBox txt_ubicacion_3 
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
         Left            =   1335
         TabIndex        =   6
         Top             =   1155
         Width           =   4110
      End
      Begin VB.TextBox txt_ubicacion_2 
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
         Left            =   1335
         TabIndex        =   4
         Top             =   720
         Width           =   4110
      End
      Begin VB.TextBox txt_ubicacion_1 
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
         Left            =   1335
         TabIndex        =   2
         Top             =   270
         Width           =   4110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sabado:"
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
         Left            =   105
         TabIndex        =   11
         Top             =   2535
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Viernes:"
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
         Left            =   90
         TabIndex        =   9
         Top             =   2085
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jueves:"
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
         Left            =   105
         TabIndex        =   7
         Top             =   1635
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Miercoles:"
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
         Left            =   90
         TabIndex        =   5
         Top             =   1185
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Martes:"
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
         TabIndex        =   3
         Top             =   750
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lunes:"
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
         Left            =   105
         TabIndex        =   1
         Top             =   300
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmoracle_ubicaciones_kanbans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    rs.Open "select * from tb_oracle_ubicaciones_motor_logistico where clave = '" + var_subinventario_kanban + "' and codigo = '" + var_codigo_kanban + "'", cnn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
       Me.txt_ubicacion_1 = IIf(IsNull(rs!ubicacion_1), "", rs!ubicacion_1)
       Me.txt_ubicacion_2 = IIf(IsNull(rs!ubicacion_2), "", rs!ubicacion_2)
       Me.txt_ubicacion_3 = IIf(IsNull(rs!ubicacion_3), "", rs!ubicacion_3)
       Me.txt_ubicacion_4 = IIf(IsNull(rs!ubicacion_4), "", rs!ubicacion_4)
       Me.txt_ubicacion_5 = IIf(IsNull(rs!ubicacion_5), "", rs!ubicacion_5)
       Me.txt_ubicacion_6 = IIf(IsNull(rs!ubicacion_6), "", rs!ubicacion_6)
    Else
       Me.txt_ubicacion_1 = ""
       Me.txt_ubicacion_2 = ""
       Me.txt_ubicacion_3 = ""
       Me.txt_ubicacion_4 = ""
       Me.txt_ubicacion_5 = ""
       Me.txt_ubicacion_6 = ""
    End If
    rs.Close

End Sub
