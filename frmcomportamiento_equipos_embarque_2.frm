VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcomportamiento_equipos_embarque_2 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gráfica de surtido"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   7005
      Left            =   5925
      TabIndex        =   28
      Top             =   -90
      Width           =   5685
      Begin MSComctlLib.ProgressBar prg_barra_equipo_8 
         Height          =   240
         Left            =   165
         TabIndex        =   29
         Top             =   510
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_9 
         Height          =   240
         Left            =   165
         TabIndex        =   30
         Top             =   1492
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_10 
         Height          =   240
         Left            =   165
         TabIndex        =   31
         Top             =   2474
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_11 
         Height          =   240
         Left            =   165
         TabIndex        =   32
         Top             =   3420
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_12 
         Height          =   240
         Left            =   165
         TabIndex        =   33
         Top             =   4455
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_13 
         Height          =   240
         Left            =   165
         TabIndex        =   34
         Top             =   5420
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_14 
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   6405
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_12 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2108
         TabIndex        =   81
         Top             =   4140
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_14 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2108
         TabIndex        =   80
         Top             =   6105
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_13 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2108
         TabIndex        =   79
         Top             =   5100
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_11 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2108
         TabIndex        =   78
         Top             =   3105
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_10 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2108
         TabIndex        =   77
         Top             =   2175
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_9 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2108
         TabIndex        =   76
         Top             =   1185
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_8 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2108
         TabIndex        =   75
         Top             =   195
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_porcentaje_equipo_8 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   135
         TabIndex        =   61
         Top             =   750
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_13 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   135
         TabIndex        =   60
         Top             =   5655
         Width           =   5250
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   180
         TabIndex        =   59
         Top             =   5040
         Width           =   1440
      End
      Begin VB.Label lbl_cantidad_total_equipo_13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4095
         TabIndex        =   58
         Top             =   5115
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_equipo_12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   57
         Top             =   4125
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_equipo_11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "1,600"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   56
         Top             =   3105
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_equipo_9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "2,550"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   55
         Top             =   1125
         Width           =   1305
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo  8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   180
         TabIndex        =   54
         Top             =   150
         Width           =   1365
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   180
         TabIndex        =   53
         Top             =   1095
         Width           =   1275
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo  10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   180
         TabIndex        =   52
         Top             =   2085
         Width           =   1530
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   180
         TabIndex        =   51
         Top             =   3075
         Width           =   1440
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   180
         TabIndex        =   50
         Top             =   4065
         Width           =   1440
      End
      Begin VB.Label lbl_cantidad_total_equipo_8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   49
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   47
         Top             =   1890
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   46
         Top             =   3345
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   45
         Top             =   4740
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   44
         Top             =   6090
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label lbl_porcentaje_equipo_10 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   135
         TabIndex        =   43
         Top             =   2730
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_11 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   135
         TabIndex        =   42
         Top             =   3675
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_12 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   135
         TabIndex        =   41
         Top             =   4695
         Width           =   5250
      End
      Begin VB.Label lbl_cantidad_total_equipo_10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "345"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   40
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label lbl_porcentaje_equipo_9 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   135
         TabIndex        =   39
         Top             =   1740
         Width           =   5250
      End
      Begin VB.Label lbl_cantidad_total_equipo_14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4050
         TabIndex        =   38
         Top             =   6105
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   360
         Left            =   135
         TabIndex        =   37
         Top             =   6015
         Width           =   1440
      End
      Begin VB.Label lbl_porcentaje_equipo_14 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   135
         TabIndex        =   36
         Top             =   6645
         Width           =   5250
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   7005
      Left            =   150
      TabIndex        =   0
      Top             =   -90
      Width           =   5685
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1905
         Top             =   0
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_1 
         Height          =   240
         Left            =   165
         TabIndex        =   1
         Top             =   510
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_2 
         Height          =   240
         Left            =   165
         TabIndex        =   2
         Top             =   1485
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_3 
         Height          =   240
         Left            =   165
         TabIndex        =   3
         Top             =   2474
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_4 
         Height          =   240
         Left            =   165
         TabIndex        =   4
         Top             =   3420
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_5 
         Height          =   240
         Left            =   165
         TabIndex        =   5
         Top             =   4438
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_6 
         Height          =   240
         Left            =   165
         TabIndex        =   6
         Top             =   5420
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prg_barra_equipo_7 
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   6405
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_5 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1995
         TabIndex        =   74
         Top             =   4140
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_7 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1995
         TabIndex        =   73
         Top             =   6105
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1995
         TabIndex        =   72
         Top             =   5100
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_4 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1995
         TabIndex        =   71
         Top             =   3105
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_3 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1995
         TabIndex        =   70
         Top             =   2175
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1995
         TabIndex        =   69
         Top             =   1185
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_ruta_equipo_1 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1995
         TabIndex        =   68
         Top             =   195
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_porcentaje_equipo_1 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   15
         TabIndex        =   67
         Top             =   705
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   15
         TabIndex        =   66
         Top             =   5655
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_3 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   15
         TabIndex        =   65
         Top             =   2730
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_4 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   15
         TabIndex        =   64
         Top             =   3675
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_5 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   15
         TabIndex        =   63
         Top             =   4695
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   15
         TabIndex        =   62
         Top             =   1740
         Width           =   5250
      End
      Begin VB.Label lbl_porcentaje_equipo_7 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   15
         TabIndex        =   27
         Top             =   6660
         Width           =   5250
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   135
         TabIndex        =   26
         Top             =   6000
         Width           =   1275
      End
      Begin VB.Label lbl_cantidad_total_equipo_7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4050
         TabIndex        =   25
         Top             =   6105
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_equipo_3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "345"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   23
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label lbl_integrantes_equipo_5 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   22
         Top             =   6090
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label lbl_integrantes_equipo_4 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   21
         Top             =   4740
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label lbl_integrantes_equipo_3 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   20
         Top             =   3345
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label lbl_integrantes_equipo_2 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   19
         Top             =   1890
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label lbl_integrantes_equipo_1 
         BackColor       =   &H80000009&
         Height          =   690
         Left            =   0
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label lbl_cantidad_total_equipo_1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "1,230"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   17
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   16
         Top             =   4050
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   15
         Top             =   3060
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo  3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   14
         Top             =   2070
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   13
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo  1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   12
         Top             =   135
         Width           =   1365
      End
      Begin VB.Label lbl_cantidad_total_equipo_2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "2,550"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   11
         Top             =   1125
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_equipo_4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "1,600"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   10
         Top             =   3105
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_equipo_5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4140
         TabIndex        =   9
         Top             =   4125
         Width           =   1305
      End
      Begin VB.Label lbl_cantidad_total_equipo_6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4095
         TabIndex        =   8
         Top             =   5115
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Equipo 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   7
         Top             =   5025
         Width           =   1275
      End
   End
   Begin VB.Label lbl_cantidad_total_surtir 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      Caption         =   "1,230"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2190
      TabIndex        =   87
      Top             =   6945
      Width           =   1230
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000009&
      Caption         =   "Piezas a surtir:"
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
      Height          =   315
      Left            =   105
      TabIndex        =   86
      Top             =   6945
      Width           =   2310
   End
   Begin VB.Label lbl_porcentaje_global 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "1,230"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   10365
      TabIndex        =   85
      Top             =   6945
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000009&
      Caption         =   "Porcentaje global:"
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
      Height          =   315
      Left            =   8055
      TabIndex        =   84
      Top             =   6945
      Width           =   2310
   End
   Begin VB.Label lbl_cantidad_total_surtida 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      Caption         =   "1,230"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5955
      TabIndex        =   83
      Top             =   6945
      Width           =   1500
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000009&
      Caption         =   "Piezas surtidas:"
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
      Height          =   315
      Left            =   3885
      TabIndex        =   82
      Top             =   6945
      Width           =   2310
   End
End
Attribute VB_Name = "frmcomportamiento_equipos_embarque_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub porcentaje()
   
   var_mes = CStr(Month(Date))
   var_dia = CStr(Day(Date))
   If Len(var_mes) = 1 Then
      var_mes = "0" + var_mes
   End If
   If Len(var_dia) = 1 Then
      var_dia = "0" + var_dia
   End If
   var_fecha_numero = CDbl(CStr(Year(Date)) + var_mes + var_dia)
   
   Dim var_cantidad_surtida As Double
   Dim var_porcentaje As Double
   Dim var_cantidad_surtir As Double
   Dim var_cantidad_total_surtida As Double
   var_fecha_numero = var_fecha_numero
   On Error GoTo salir:
   rs.Open "SELECT * FROM VW_DETALLE_EQUIPOS_ORDEN_SURTIDO WHERE INTE_EQU_NUMERO = " + CStr(var_fecha_numero), cnn, adOpenDynamic, adLockOptimistic
   var_cantidad_total_surtida = 0
   While Not rs.EOF
         If rs!inte_equ_equipo = 1 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_1)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_porcentaje = 100
               var_cantidad_surtida = var_cantidad_surtir
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_1 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_1.Value = var_porcentaje
            Me.prg_barra_equipo_1.Refresh
         End If
         If rs!inte_equ_equipo = 2 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_2)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_porcentaje = 100
               var_cantidad_surtida = var_cantidad_surtir
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_2 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_2.Value = var_porcentaje
            Me.prg_barra_equipo_2.Refresh
         End If
         If rs!inte_equ_equipo = 3 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_3)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_3 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_3.Value = var_porcentaje
            Me.prg_barra_equipo_3.Refresh
         End If
         If rs!inte_equ_equipo = 4 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_4)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_porcentaje = 100
               var_cantidad_surtida = var_cantidad_surtir
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_4 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_4.Value = var_porcentaje
            Me.prg_barra_equipo_4.Refresh
         End If
         If rs!inte_equ_equipo = 5 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_5)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_5 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_5.Value = var_porcentaje
            Me.prg_barra_equipo_5.Refresh
         End If
         If rs!inte_equ_equipo = 6 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_6)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_6 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_6.Value = var_porcentaje
            Me.prg_barra_equipo_6.Refresh
         End If
         
         
         If rs!inte_equ_equipo = 7 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_7)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_7 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_7.Value = var_porcentaje
            Me.prg_barra_equipo_7.Refresh
         End If
         
         If rs!inte_equ_equipo = 8 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_8)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_8 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_8.Value = var_porcentaje
            Me.prg_barra_equipo_8.Refresh
         End If
         
         If rs!inte_equ_equipo = 9 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_9)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_9 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_9.Value = var_porcentaje
            Me.prg_barra_equipo_9.Refresh
         End If
         
         
         If rs!inte_equ_equipo = 10 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_10)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_10 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_10.Value = var_porcentaje
            Me.prg_barra_equipo_10.Refresh
         End If
         
         If rs!inte_equ_equipo = 11 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_11)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_11 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_11.Value = var_porcentaje
            Me.prg_barra_equipo_11.Refresh
         End If
         
         If rs!inte_equ_equipo = 12 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_12)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_12 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_12.Value = var_porcentaje
            Me.prg_barra_equipo_12.Refresh
         End If
         
         If rs!inte_equ_equipo = 13 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_13)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_13 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_13.Value = var_porcentaje
            Me.prg_barra_equipo_13.Refresh
         End If
         
         
         If rs!inte_equ_equipo = 14 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_14)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_14 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_14.Value = var_porcentaje
            Me.prg_barra_equipo_14.Refresh
         End If
         
         rs.MoveNext
   Wend
   rs.Close
   lbl_cantidad_total_surtida = Format(var_cantidad_total_surtida, "###,###,##0")
   var_cantidad_surtida = CDbl(lbl_cantidad_total_surtida)
   var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_surtir)
   If var_cantidad_surtir > 0 Then
      var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
   Else
      var_porcentaje = 0
   End If
   If var_porcentaje > 100 Then
      var_porcentaje = 100
   End If
   
   Me.lbl_porcentaje_global = CStr(Round(var_porcentaje, 2)) + " %"
   Exit Sub
salir:
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

End Sub
Private Sub equipos()
   Dim var_contador As Integer
   Dim var_ruta As String
   Dim var_cantidad As String
   Dim var_cantidad_total As Double
   var_mes = CStr(Month(Date))
   var_dia = CStr(Day(Date))
   If Len(var_mes) = 1 Then
      var_mes = "0" + var_mes
   End If
   If Len(var_dia) = 1 Then
      var_dia = "0" + var_dia
   End If
   var_fecha_numero = CDbl(CStr(Year(Date)) + var_mes + var_dia)
   'Me.lst_rutas_equipo_1.Clear
   'Me.lst_rutas_equipo_2.Clear
   'Me.lst_rutas_equipo_3.Clear
   'Me.lst_rutas_equipo_4.Clear
   'Me.lst_rutas_equipo_5.Clear
   'Me.lst_rutas_equipo_6.Clear
   'Me.lst_rutas_equipo_7.Clear
   'Me.lst_rutas_equipo_8.Clear
   'Me.lst_rutas_equipo_9.Clear
   'Me.lst_rutas_equipo_10.Clear
   'Me.lst_rutas_equipo_11.Clear
   'Me.lst_rutas_equipo_12.Clear
   'Me.lst_rutas_equipo_13.Clear
   'Me.lst_rutas_equipo_14.Clear
   cantidad_total_rutas = 0
   For var_j = 1 To 14
      var_equipo = var_j
      If var_j = 1 Then
         
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = 1", cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         var_contador = 1
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_1.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_1 = Format(var_cantidad_total, "###,###,##0")
         
         
         
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            'While Not rs.EOF
            '      rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            '      If Not rsaux.EOF Then
            '         If Trim(Me.lbl_integrantes_equipo_1) = "" Then
            '            Me.lbl_integrantes_equipo_1 = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            '         Else
            '            Me.lbl_integrantes_equipo_1 = Me.lbl_integrantes_equipo_1 + ", " + IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            '         End If
            '      End If
            '      rsaux.Close
            '      rs.MoveNext
            'Wend
         Else
            Me.lbl_porcentaje_equipo_1 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_1 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_1 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_1) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 2 Then
         
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = 2", cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_2.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_2 = Format(var_cantidad_total, "###,###,##0")
         
         
         
         
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            'While Not rs.EOF
                  'rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  'If Not rsaux.EOF Then
                  '   If Trim(Me.lbl_integrantes_equipo_2) = "" Then
                  '      Me.lbl_integrantes_equipo_2 = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
                  '   Else
                  '      Me.lbl_integrantes_equipo_2 = Me.lbl_integrantes_equipo_2 + ", " + IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
                  '   End If
                  'End If
                  'rsaux.Close
                  rs.MoveNext
            'Wend
         Else
            Me.lbl_porcentaje_equipo_2 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_2 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_2 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_2) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 3 Then
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = 3", cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_3.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_3 = Format(var_cantidad_total, "###,###,##0")
         
         
         
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
         '   While Not rs.EOF
         '         rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
         '         If Not rsaux.EOF Then
         '            If Trim(Me.lbl_integrantes_equipo_3) = "" Then
         '               Me.lbl_integrantes_equipo_3 = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
         '            Else
         '               Me.lbl_integrantes_equipo_3 = Me.lbl_integrantes_equipo_3 + ", " + IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
         '            End If
         '         End If
         '         rsaux.Close
         '         rs.MoveNext
         '   Wend
         Else
            Me.lbl_porcentaje_equipo_3 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_3 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_3 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_3) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 4 Then
         
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = 4", cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_4.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_4 = Format(var_cantidad_total, "###,###,##0")
         
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            'While Not rs.EOF
            '      rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            '      If Not rsaux.EOF Then
            '         If Trim(Me.lbl_integrantes_equipo_4) = "" Then
            '            Me.lbl_integrantes_equipo_4 = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            '         Else
            '            Me.lbl_integrantes_equipo_4 = Me.lbl_integrantes_equipo_4 + ", " + IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            '         End If
            '      End If
            '      rsaux.Close
            '      rs.MoveNext
            'Wend
         Else
            Me.lbl_porcentaje_equipo_4 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_4 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_4 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_4) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 5 Then
         
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = 5", cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_5.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_5 = Format(var_cantidad_total, "###,###,##0")
         
         
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            'While Not rs.EOF
            '      rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            '      If Not rsaux.EOF Then
            '         If Trim(Me.lbl_integrantes_equipo_5) = "" Then
            '            Me.lbl_integrantes_equipo_5 = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            '         Else
            '            Me.lbl_integrantes_equipo_5 = Me.lbl_integrantes_equipo_5 + ", " + IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            '         End If
            '      End If
            '      rsaux.Close
            '      rs.MoveNext
            'Wend
         Else
            Me.lbl_porcentaje_equipo_5 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_5 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_5 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_5) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
''' equipo 6
   
   
   
      If var_j = 6 Then
         
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = 6", cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_6.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_6 = Format(var_cantidad_total, "###,###,##0")
         
         
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            'While Not rs.EOF
            '      rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
            '      If Not rsaux.EOF Then
            '         If Trim(Me.lbl_integrantes_equipo_5) = "" Then
            '            Me.lbl_integrantes_equipo_6 = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            '         Else
            '            Me.lbl_integrantes_equipo_6 = Me.lbl_integrantes_equipo_5 + ", " + IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
            '         End If
            '      End If
            '      rsaux.Close
            '      rs.MoveNext
            'Wend
         Else
            Me.lbl_porcentaje_equipo_6 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_6 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_6 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_6) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
   
''' fin de equipo 6
   
      If var_j = 7 Then
         
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = 7", cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_7.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_7 = Format(var_cantidad_total, "###,###,##0")
         
         
         
         
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            'While Not rs.EOF
                  'rsaux.Open "select * from tb_personal where vcha_per_personal_id = '" + rs!vcha_per_personal_id + "'", cnn, adOpenDynamic, adLockOptimistic
                  'If Not rsaux.EOF Then
                  '   If Trim(Me.lbl_integrantes_equipo_2) = "" Then
                  '      Me.lbl_integrantes_equipo_2 = IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
                  '   Else
                  '      Me.lbl_integrantes_equipo_2 = Me.lbl_integrantes_equipo_2 + ", " + IIf(IsNull(rsaux!vcha_per_nombre), "", rsaux!vcha_per_nombre)
                  '   End If
                  'End If
                  'rsaux.Close
                  rs.MoveNext
            'Wend
         Else
            Me.lbl_porcentaje_equipo_7 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_7 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_7 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_7) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
   
      If var_j = 8 Then
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_j), cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_8.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_8 = Format(var_cantidad_total, "###,###,##0")
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            Me.lbl_porcentaje_equipo_8 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_8 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_8 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_8) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 9 Then
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_j), cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_9.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_9 = Format(var_cantidad_total, "###,###,##0")
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            Me.lbl_porcentaje_equipo_9 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_9 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_9 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_9) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
      
      If var_j = 10 Then
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_j), cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_10.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_10 = Format(var_cantidad_total, "###,###,##0")
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            Me.lbl_porcentaje_equipo_10 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_10 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_10 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_10) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 11 Then
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_j), cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_11.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_11 = Format(var_cantidad_total, "###,###,##0")
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            Me.lbl_porcentaje_equipo_11 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_11 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_11 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_11) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 12 Then
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_j), cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_12.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_12 = Format(var_cantidad_total, "###,###,##0")
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            Me.lbl_porcentaje_equipo_12 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_12 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_12 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_12) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 13 Then
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_j), cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_13.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_13 = Format(var_cantidad_total, "###,###,##0")
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            Me.lbl_porcentaje_equipo_13 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_13 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_13 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_13) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   
      If var_j = 14 Then
         rs.Open "select * from vw_detalle_equipos_rutas where INTE_EQU_NUMERO = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_j), cnn, adOpenDynamic, adLockOptimistic
         var_cantidad_total = 0
         While Not rs.EOF
               var_ruta = IIf(IsNull(rs!vcha_rut_nombre), "", rs!vcha_rut_nombre)
               For var_z = Len(var_ruta) To 19
                   var_ruta = var_ruta + " "
               Next var_z
               var_cantidad = Format(IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIR), 0, rs!FLOA_ORS_CANTIDAD_SURTIR), "###,###,##0")
               For var_z = Len(var_cantidad) To 12
                   var_cantidad = " " + var_cantidad
               Next var_z
               var_ruta = var_ruta + var_cantidad
               'lst_rutas_equipo_14.AddItem (var_ruta)
               var_cantidad_total = var_cantidad_total + CDbl(var_cantidad)
               cantidad_total_rutas = cantidad_total_rutas + CDbl(var_cantidad)
               rs.MoveNext
         Wend
         rs.Close
         Me.lbl_cantidad_total_ruta_equipo_14 = Format(var_cantidad_total, "###,###,##0")
         rs.Open "select * from tb_detalle_equipos_personal where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If rs.EOF Then
            Me.lbl_porcentaje_equipo_14 = "NO CREADO"
         End If
         rs.Close
         Me.lbl_cantidad_total_equipo_14 = "0"
         rs.Open "select * from tb_detalle_equipos_orden_surtido where inte_equ_numero = " + CStr(var_fecha_numero) + " and inte_equ_equipo = " + CStr(var_equipo), cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            While Not rs.EOF
                  Me.lbl_cantidad_total_equipo_14 = Format(CStr(CDbl(Me.lbl_cantidad_total_equipo_14) + IIf(IsNull(rs!floa_ors_cantidad), 0, rs!floa_ors_cantidad)), "###,###,##0")
                  rs.MoveNext
            Wend
         End If
         rs.Close
      End If
   

   
   
   Next var_j
   Me.lbl_cantidad_total_surtir = Format(cantidad_total_rutas, "###,###,##0")
   Call porcentaje
End Sub



Private Sub Form_Load()
   Top = 0
   Left = 0
   lbl_fecha = Format(Date, "Long Date")
   Me.lbl_cantidad_total_equipo_1 = "0"
   Me.lbl_cantidad_total_equipo_2 = "0"
   Me.lbl_cantidad_total_equipo_3 = "0"
   Me.lbl_cantidad_total_equipo_4 = "0"
   Me.lbl_cantidad_total_equipo_5 = "0"
   Me.lbl_cantidad_total_equipo_6 = "0"
   Me.lbl_cantidad_total_equipo_7 = "0"
   Me.lbl_cantidad_total_equipo_8 = "0"
   Me.lbl_cantidad_total_equipo_9 = "0"
   Me.lbl_cantidad_total_equipo_10 = "0"
   Me.lbl_cantidad_total_equipo_11 = "0"
   Me.lbl_cantidad_total_equipo_12 = "0"
   Me.lbl_cantidad_total_equipo_13 = "0"
   Me.lbl_cantidad_total_equipo_14 = "0"
   Me.lbl_cantidad_total_ruta_equipo_1 = "0"
   Me.lbl_cantidad_total_ruta_equipo_2 = "0"
   Me.lbl_cantidad_total_ruta_equipo_3 = "0"
   Me.lbl_cantidad_total_ruta_equipo_4 = "0"
   Me.lbl_cantidad_total_ruta_equipo_5 = "0"
   Me.lbl_cantidad_total_ruta_equipo_6 = "0"
   Me.lbl_cantidad_total_ruta_equipo_7 = "0"
   Me.lbl_cantidad_total_ruta_equipo_8 = "0"
   Me.lbl_cantidad_total_ruta_equipo_9 = "0"
   Me.lbl_cantidad_total_ruta_equipo_10 = "0"
   Me.lbl_cantidad_total_ruta_equipo_11 = "0"
   Me.lbl_cantidad_total_ruta_equipo_12 = "0"
   Me.lbl_cantidad_total_ruta_equipo_13 = "0"
   Me.lbl_cantidad_total_ruta_equipo_14 = "0"
   Me.lbl_cantidad_total_surtida = "0"
   Me.lbl_cantidad_total_surtir = "0"
   Me.lbl_porcentaje_equipo_1 = "0%"
   Me.lbl_porcentaje_equipo_2 = "0%"
   Me.lbl_porcentaje_equipo_3 = "0%"
   Me.lbl_porcentaje_equipo_4 = "0%"
   Me.lbl_porcentaje_equipo_5 = "0%"
   Me.lbl_porcentaje_equipo_6 = "0%"
   Me.lbl_porcentaje_equipo_7 = "0%"
   Me.lbl_porcentaje_equipo_8 = "0%"
   Me.lbl_porcentaje_equipo_9 = "0%"
   Me.lbl_porcentaje_equipo_10 = "0%"
   Me.lbl_porcentaje_equipo_11 = "0%"
   Me.lbl_porcentaje_equipo_12 = "0%"
   Me.lbl_porcentaje_equipo_13 = "0%"
   Me.lbl_porcentaje_equipo_14 = "0%"
   Me.lbl_porcentaje_global = "0%"
   equipos
   Me.Timer1.Enabled = True
End Sub

Private Sub Label9_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub Timer1_Timer()
   var_mes = CStr(Month(Date))
   var_dia = CStr(Day(Date))
   If Len(var_mes) = 1 Then
      var_mes = "0" + var_mes
   End If
   If Len(var_dia) = 1 Then
      var_dia = "0" + var_dia
   End If
   var_fecha_numero = CDbl(CStr(Year(Date)) + var_mes + var_dia)
   
   Dim var_cantidad_surtida As Double
   Dim var_porcentaje As Double
   Dim var_cantidad_surtir As Double
   Dim var_cantidad_total_surtida As Double
   var_fecha_numero = var_fecha_numero
   On Error GoTo salir:
   rs.Open "SELECT * FROM VW_DETALLE_EQUIPOS_ORDEN_SURTIDO WHERE INTE_EQU_NUMERO = " + CStr(var_fecha_numero), cnn, adOpenDynamic, adLockOptimistic
   var_cantidad_total_surtida = 0
   While Not rs.EOF
         If rs!inte_equ_equipo = 1 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_1)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_porcentaje = 100
               var_cantidad_surtida = var_cantidad_surtir
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_1 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_1.Value = var_porcentaje
            Me.prg_barra_equipo_1.Refresh
         End If
         If rs!inte_equ_equipo = 2 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_2)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_porcentaje = 100
               var_cantidad_surtida = var_cantidad_surtir
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_2 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_2.Value = var_porcentaje
            Me.prg_barra_equipo_2.Refresh
         End If
         If rs!inte_equ_equipo = 3 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_3)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_3 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_3.Value = var_porcentaje
            Me.prg_barra_equipo_3.Refresh
         End If
         If rs!inte_equ_equipo = 4 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_4)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_porcentaje = 100
               var_cantidad_surtida = var_cantidad_surtir
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_4 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_4.Value = var_porcentaje
            Me.prg_barra_equipo_4.Refresh
         End If
         If rs!inte_equ_equipo = 5 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_5)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_5 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_5.Value = var_porcentaje
            Me.prg_barra_equipo_5.Refresh
         End If
         If rs!inte_equ_equipo = 6 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_6)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_6 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_6.Value = var_porcentaje
            Me.prg_barra_equipo_6.Refresh
         End If
         If rs!inte_equ_equipo = 7 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_7)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_7 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_7.Value = var_porcentaje
            Me.prg_barra_equipo_7.Refresh
         End If
         If rs!inte_equ_equipo = 8 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_8)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_8 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_8.Value = var_porcentaje
            Me.prg_barra_equipo_8.Refresh
         End If
         If rs!inte_equ_equipo = 9 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_9)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_9 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_9.Value = var_porcentaje
            Me.prg_barra_equipo_9.Refresh
         End If
         If rs!inte_equ_equipo = 10 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_10)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_10 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_10.Value = var_porcentaje
            Me.prg_barra_equipo_10.Refresh
         End If
         If rs!inte_equ_equipo = 11 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_11)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_11 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_11.Value = var_porcentaje
            Me.prg_barra_equipo_11.Refresh
         End If
         If rs!inte_equ_equipo = 12 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_12)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_12 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_12.Value = var_porcentaje
            Me.prg_barra_equipo_12.Refresh
         End If
         If rs!inte_equ_equipo = 13 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_13)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_13 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_13.Value = var_porcentaje
            Me.prg_barra_equipo_13.Refresh
         End If
         If rs!inte_equ_equipo = 14 Then
            var_cantidad_surtida = IIf(IsNull(rs!FLOA_ORS_CANTIDAD_SURTIDA), 0, rs!FLOA_ORS_CANTIDAD_SURTIDA)
            var_cantidad_total_surtida = var_cantidad_total_surtida + var_cantidad_surtida
            var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_equipo_14)
            var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
            If var_porcentaje > 100 Then
               var_cantidad_surtida = var_cantidad_surtir
               var_porcentaje = 100
            End If
            var_cantidad_falta = var_cantidad_surtir - var_cantidad_surtida
            Me.lbl_porcentaje_equipo_14 = CStr(var_cantidad_surtida) + " Piezas   " + CStr(Round(var_porcentaje, 2)) + " %   faltan " + CStr(var_cantidad_falta) + " Piezas"
            Me.prg_barra_equipo_14.Value = var_porcentaje
            Me.prg_barra_equipo_14.Refresh
         End If
         
         rs.MoveNext
   Wend
   rs.Close
   lbl_cantidad_total_surtida = Format(var_cantidad_total_surtida, "###,###,##0")
   var_cantidad_surtida = CDbl(lbl_cantidad_total_surtida)
   var_cantidad_surtir = CDbl(Me.lbl_cantidad_total_surtir)
   If var_cantidad_surtir > 0 Then
      var_porcentaje = (var_cantidad_surtida * 100) / var_cantidad_surtir
   Else
      var_porcentaje = 0
   End If
   If var_porcentaje > 100 Then
      var_porcentaje = 100
   End If
   
   Me.lbl_porcentaje_global = CStr(Round(var_porcentaje, 2)) + " %"
   Exit Sub
salir:
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
End Sub
