VERSION 5.00
Begin VB.Form frmConfiguracionTamaños 
   Caption         =   "Configuracion Tamaños Etiquetas Precios"
   ClientHeight    =   10725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   10725
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Campo de Base de Datos"
      Height          =   1335
      Left            =   120
      TabIndex        =   94
      Top             =   9120
      Width           =   7215
      Begin VB.TextBox txt_CampoEncabezado1 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txt_CampoEncabezado2 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label33 
         Caption         =   "Encabezado 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label32 
         Caption         =   "Encabezado 2:"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "Precio:"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   2760
         TabIndex        =   103
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Medidas:"
         Height          =   255
         Left            =   2760
         TabIndex        =   102
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Color:"
         Height          =   255
         Left            =   2760
         TabIndex        =   101
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Localizacion:"
         Height          =   255
         Left            =   5280
         TabIndex        =   100
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmd_nuevo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmConfiguracionTamaños.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   88
      ToolTipText     =   "Nuevo Alt + N"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      Picture         =   "frmConfiguracionTamaños.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   470
      Picture         =   "frmConfiguracionTamaños.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   85
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   120
      TabIndex        =   86
      Top             =   270
      Width           =   7245
   End
   Begin VB.Frame fraEspacioEntreLineas 
      Caption         =   "Espacios Entre Lineas"
      Height          =   1335
      Left            =   120
      TabIndex        =   68
      Top             =   7680
      Width           =   7215
      Begin VB.TextBox txt_tamLineaLoc 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_tamLineaColor 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_tamLineaMedia 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt_tamLineaCodigo 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_tamLineaPrecio 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt_tamLineaEnc2 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_tamLineaEnc1 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Localizacion:"
         Height          =   255
         Left            =   3600
         TabIndex        =   82
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "Color:"
         Height          =   255
         Left            =   1920
         TabIndex        =   81
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Medidas:"
         Height          =   255
         Left            =   1920
         TabIndex        =   80
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   1920
         TabIndex        =   79
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "Precio:"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Encabezado 2:"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Encabezado 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FraCaracteres 
      Caption         =   "Caracteres por Linea"
      Height          =   1335
      Left            =   120
      TabIndex        =   53
      Top             =   6240
      Width           =   7215
      Begin VB.TextBox txt_CaracteresEnc1 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_CaracteresEnc2 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_CaracteresPrecio 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt_CaracteresCodigo 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_CaracteresMedidas 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt_CaracteresColor 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_CaracteresLoc 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "Encabezado 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Encabezado 2:"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Precio:"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   1920
         TabIndex        =   64
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Medidas:"
         Height          =   255
         Left            =   1920
         TabIndex        =   63
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Color:"
         Height          =   255
         Left            =   1920
         TabIndex        =   62
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Localizacion:"
         Height          =   255
         Left            =   3600
         TabIndex        =   61
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FraUbicacion 
      BackColor       =   &H8000000A&
      Caption         =   "Ubicacion"
      Height          =   1935
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   7215
      Begin VB.TextBox txt_Posi_x_SignoPrecio 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_y_SignoPrecio 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   1560
         Width           =   495
      End
      Begin VB.ComboBox cbo_rotacion 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmConfiguracionTamaños.frx":083E
         Left            =   5400
         List            =   "frmConfiguracionTamaños.frx":0848
         TabIndex        =   52
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txt_Posi_y_Medidas 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_x_Medidas 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_y_Color 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_x_Color 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_y_Codigo 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_x_Codigo 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_y_Loc 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_x_Loc 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_y_Precio 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_y_Enc2 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_y_Enc1 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_x_Precio 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_x_Enc2 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txt_Posi_x_Enc1 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "$:"
         Height          =   255
         Left            =   2520
         TabIndex        =   91
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbl_rotacion 
         Caption         =   "Rotacion:"
         Height          =   255
         Left            =   4560
         TabIndex        =   51
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Y"
         Height          =   255
         Left            =   3960
         TabIndex        =   50
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "X"
         Height          =   255
         Left            =   3360
         TabIndex        =   49
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Y"
         Height          =   255
         Left            =   1920
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "X"
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Localizacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Color:"
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Medidas:"
         Height          =   255
         Left            =   2520
         TabIndex        =   30
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Precio:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Encabezado 2:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Encabezado 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame FraFuentes 
      Caption         =   "Fuentes"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   7215
      Begin VB.TextBox txt_fuenteSignoPrecio 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_fuenteUbicacion 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_fuentecolor 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_fuenteMedidas 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt_fuenteCodigo 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_fuentePrecio 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt_fuenteEnc2 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_FuenteEnc1 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbl_fuenteSignoPrecio 
         Caption         =   "$:"
         Height          =   255
         Left            =   3600
         TabIndex        =   90
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbl_fuenteLocalizacion 
         Caption         =   "Localizacion:"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbl_fuenteColor 
         Caption         =   "Color:"
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbl_FuenteMedidas 
         Caption         =   "Medidas:"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl_fuenteCodigo 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbl_fuentePrecio 
         Caption         =   "Precio:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl_fuenteEnc2 
         Caption         =   "Encabezado 2:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbl_FuenteEnc1 
         Caption         =   "Encabezado 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraMedidas 
      Caption         =   "Medidas"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   7215
      Begin VB.TextBox txt_lineaNegra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   87
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txt_largo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt_ancho 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_TamLineaNegra 
         Caption         =   "Linea Negra"
         Height          =   255
         Left            =   2040
         TabIndex        =   83
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_largo 
         Caption         =   "Largo:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl_ancho 
         Caption         =   "Ancho:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame frmTamaños 
      Caption         =   "Tamaños"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin VB.TextBox txt_descripcion 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl_enc_nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lbl_Codigo 
         Caption         =   "ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmConfiguracionTamaños"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnnSID As New ADODB.Connection
Private Function fun_valida_Informacion()
    Dim str_valida As String
    fun_valida_Informacion = str_valida
    
    If txt_tamLineaCodigo.Text = "" Then
        txt_tamLineaCodigo.Text = -1
    Else
        If Not IsNumeric(txt_tamLineaCodigo.Text) Then
            str_valida = "El espacio entre linea de Codigo debe ser numerico "
            txt_tamLineaCodigo.SetFocus
        End If
    End If
    
    If txt_tamLineaColor.Text = "" Then
        txt_tamLineaColor.Text = -1
    Else
        If Not IsNumeric(txt_tamLineaColor.Text) Then
            str_valida = str_valida & vbCrLf & "El espacio entre linea de ''Color'' debe ser numerico "
            txt_tamLineaColor.SetFocus
        End If
    End If
    
    If txt_tamLineaEnc1.Text = "" Then
        txt_tamLineaEnc1.Text = -1
    Else
        If Not IsNumeric(txt_tamLineaEnc1.Text) Then
            str_valida = str_valida & vbCrLf & "El espacio entre linea del Encabezado1 debe ser numerico "
            txt_tamLineaEnc1.SetFocus
        End If
    End If
    If txt_tamLineaEnc2.Text = "" Then
        txt_tamLineaEnc2.Text = -1
    Else
        If Not IsNumeric(txt_tamLineaEnc2.Text) Then
            str_valida = str_valida & vbCrLf & "El espacio entre linea del Encabezado2 debe ser numerico "
            txt_tamLineaEnc2.SetFocus
        End If
    End If
    
    If txt_tamLineaLoc.Text = "" Then
        txt_tamLineaLoc.Text = -1
    Else
        If Not IsNumeric(txt_tamLineaLoc.Text) Then
            str_valida = str_valida & vbCrLf & "El espacio entre linea de Localizacion debe ser numerico "
            txt_tamLineaLoc.SetFocus
        End If
    End If
    
    If txt_tamLineaMedia.Text = "" Then
        txt_tamLineaMedia.Text = -1
    Else
        If Not IsNumeric(txt_tamLineaMedia.Text) Then
            str_valida = str_valida & vbCrLf & "El espacio entre linea de Medida debe ser numerico "
            txt_tamLineaMedia.SetFocus
        End If
    End If
    
    If txt_tamLineaPrecio.Text = "" Then
        txt_tamLineaPrecio.Text = -1
    Else
        If Not IsNumeric(txt_tamLineaPrecio.Text) Then
            str_valida = str_valida & vbCrLf & "El espacio entre linea de Precio debe ser numerico "
            txt_tamLineaPrecio.SetFocus
        End If
    End If
    
    
    If txt_CaracteresCodigo.Text = "" Then
        txt_CaracteresCodigo.Text = -1
    Else
        If Not IsNumeric(txt_CaracteresCodigo.Text) Then
            str_valida = str_valida & vbCrLf & "Los caracteres por linea de Codigo debe ser numerico "
            txt_CaracteresCodigo.SetFocus
        End If
    End If
    
    If txt_CaracteresColor.Text = "" Then
        txt_CaracteresColor.Text = -1
    Else
        If Not IsNumeric(txt_CaracteresColor.Text) Then
            str_valida = str_valida & vbCrLf & "Los caracteres por linea de Color debe ser numerico "
            txt_CaracteresColor.SetFocus
        End If
    End If
    If txt_CaracteresEnc1.Text = "" Then
        txt_CaracteresEnc1.Text = -1
    Else
        If Not IsNumeric(txt_CaracteresEnc1.Text) Then
            str_valida = str_valida & vbCrLf & "Los caracteres por linea de Encabezado1 debe ser numerico "
            txt_CaracteresEnc1.SetFocus
        End If
    End If
    If txt_CaracteresEnc2.Text = "" Then
        txt_CaracteresEnc2.Text = -1
    Else
        If Not IsNumeric(txt_CaracteresEnc2.Text) Then
            str_valida = str_valida & vbCrLf & "Los caracteres por linea de Encabezado2 debe ser numerico "
            txt_CaracteresEnc2.SetFocus
        End If
    End If
    If txt_CaracteresLoc.Text = "" Then
        txt_CaracteresLoc.Text = -1
    Else
        If Not IsNumeric(txt_CaracteresLoc.Text) Then
            str_valida = str_valida & vbCrLf & "Los caracteres por linea de Localizacion debe ser numerico "
            txt_CaracteresLoc.SetFocus
        End If
    End If
    If txt_CaracteresMedidas.Text = "" Then
        txt_CaracteresMedidas.Text = -1
    Else
        If Not IsNumeric(txt_CaracteresMedidas.Text) Then
            str_valida = str_valida & vbCrLf & "Los caracteres por linea de Medida debe ser numerico "
            txt_CaracteresMedidas.SetFocus
        End If
    End If
    If txt_CaracteresPrecio.Text = "" Then
        txt_CaracteresPrecio.Text = -1
    Else
        If Not IsNumeric(txt_CaracteresPrecio.Text) Then
            str_valida = str_valida & vbCrLf & "Los caracteres por linea de Precio debe ser numerico "
            txt_CaracteresPrecio.SetFocus
        End If
    End If
    
    If txt_Posi_x_Codigo.Text = "" Then
        txt_Posi_x_Codigo.Text = -1
    Else
        If Not IsNumeric(txt_Posi_x_Codigo.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion X de Codigo debe ser numerico "
            txt_Posi_x_Codigo.SetFocus
        End If
    End If
    
    If txt_Posi_x_Color.Text = "" Then
        txt_Posi_x_Color.Text = -1
    Else
        If Not IsNumeric(txt_Posi_x_Color.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion X de Color debe ser numerico "
            txt_Posi_x_Color.SetFocus
        End If
    End If
    If txt_Posi_x_Enc1.Text = "" Then
        txt_Posi_x_Enc1.Text = -1
    Else
        If Not IsNumeric(txt_Posi_x_Enc1.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion X de Encabezado 1 debe ser numerico "
            txt_Posi_x_Enc1.SetFocus
        End If
    End If
    If txt_Posi_x_Enc2.Text = "" Then
        txt_Posi_x_Enc2.Text = -1
    Else
        If Not IsNumeric(txt_Posi_x_Enc2.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion X de Encabezado 2 debe ser numerico "
            txt_Posi_x_Enc1.SetFocus
        End If
    End If
    If txt_Posi_x_Loc.Text = "" Then
        txt_Posi_x_Loc.Text = -1
    Else
        If Not IsNumeric(txt_Posi_x_Loc.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion X de Localizacion debe ser numerico "
            txt_Posi_x_Loc.SetFocus
        End If
    End If
    If txt_Posi_x_Medidas.Text = "" Then
        txt_Posi_x_Medidas.Text = -1
    Else
        If Not IsNumeric(txt_Posi_x_Medidas.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion X de Medidas debe ser numerico "
            txt_Posi_x_Medidas.SetFocus
        End If
    End If
    If txt_Posi_x_Precio.Text = "" Then
        txt_Posi_x_Precio.Text = -1
    Else
        If Not IsNumeric(txt_Posi_x_Precio.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion X de Precio debe ser numerico "
            txt_Posi_x_Precio.SetFocus
        End If
    End If
    If txt_Posi_y_Codigo.Text = "" Then
        txt_Posi_y_Codigo.Text = -1
    Else
        If Not IsNumeric(txt_Posi_y_Codigo.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion Y de Codigo debe ser numerico "
            txt_Posi_y_Codigo.SetFocus
        End If
    End If
    
    If txt_Posi_y_Color.Text = "" Then
        txt_Posi_y_Color.Text = -1
    Else
        If Not IsNumeric(txt_Posi_y_Color.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion Y de Color debe ser numerico "
            txt_Posi_y_Color.SetFocus
        End If
    End If
    If txt_Posi_y_Enc1.Text = "" Then
        txt_Posi_y_Enc1.Text = -1
    Else
        If Not IsNumeric(txt_Posi_y_Enc1.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion Y de Encabezado1 debe ser numerico "
            txt_Posi_y_Enc1.SetFocus
        End If
    End If
    If txt_Posi_y_Enc2.Text = "" Then
        txt_Posi_y_Enc2.Text = -1
    Else
        If Not IsNumeric(txt_Posi_y_Enc2.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion Y de Encabezado2 debe ser numerico "
            txt_Posi_y_Enc1.SetFocus
        End If
    End If
    If txt_Posi_y_Loc.Text = "" Then
        txt_Posi_y_Loc.Text = -1
    Else
        If Not IsNumeric(txt_Posi_y_Loc.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion Y de Localizacion debe ser numerico "
            txt_Posi_y_Loc.SetFocus
        End If
    End If
    If txt_Posi_y_Medidas.Text = "" Then
        txt_Posi_y_Medidas.Text = -1
    Else
        If Not IsNumeric(txt_Posi_y_Medidas.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion Y de Medidas debe ser numerico "
            txt_Posi_y_Medidas.SetFocus
        End If
    End If
    If txt_Posi_y_Precio.Text = "" Then
        txt_Posi_y_Precio.Text = -1
    Else
        If Not IsNumeric(txt_Posi_y_Precio.Text) Then
            str_valida = str_valida & vbCrLf & "La posicion Y de Precio debe ser numerico "
            txt_Posi_y_Precio.SetFocus
        End If
    End If
    If Not cbo_rotacion.Text = "Sin Rotacion" Or Not cbo_rotacion.Text = "90 Grados" Then
        str_valida = str_valida & vbCrLf & "Favor de selecconar la rotacion valida. "
    End If
    
    
    If IsNumeric(txt_fuenteCodigo.Text) Then
        If txt_fuenteCodigo.Text > 5 Then
            str_valida = str_valida & vbCrLf & "La fuente de Codigo debe ser letra(a-z) o un numero menor a 6."
            txt_fuenteCodigo.SetFocus
        End If
    Else
        If Len(txt_fuenteCodigo.Text) > 1 Then
            str_valida = str_valida & vbCrLf & "Solo se acepta una letra para fuente de Codigo"
            txt_fuenteCodigo.SetFocus
        End If
    End If
    
    If IsNumeric(txt_fuentecolor.Text) Then
        If txt_fuentecolor.Text > 5 Then
            str_valida = str_valida & vbCrLf & "La fuente de Color debe ser letra(a-z) o un numero menor a 6."
            txt_fuentecolor.SetFocus
        End If
    Else
        If Len(txt_fuentecolor.Text) > 1 Then
            str_valida = str_valida & vbCrLf & "Solo se acepta una letra para fuente de Color"
            txt_fuentecolor.SetFocus
        End If
    End If
    If IsNumeric(txt_FuenteEnc1.Text) Then
        If txt_FuenteEnc1.Text > 5 Then
            str_valida = str_valida & vbCrLf & "La fuente de Encabezador1 debe ser letra(a-z) o un numero menor a 6."
            txt_FuenteEnc1.SetFocus
        End If
    Else
        If Len(txt_FuenteEnc1.Text) > 1 Then
            str_valida = str_valida & vbCrLf & "Solo se acepta una letra para fuente de Encabezado1"
            txt_FuenteEnc1.SetFocus
        End If
    End If
    If IsNumeric(txt_fuenteEnc2.Text) Then
        If txt_fuenteEnc2.Text > 5 Then
            str_valida = str_valida & vbCrLf & "La fuente de Encaebezado2 debe ser letra(a-z) o un numero menor a 6."
            txt_fuenteEnc2.SetFocus
        End If
    Else
        If Len(txt_fuenteEnc2.Text) > 1 Then
            str_valida = str_valida & vbCrLf & "Solo se acepta una letra para fuente de Encabezado2"
            txt_fuenteEnc2.SetFocus
        End If
    End If
    If IsNumeric(txt_fuenteMedidas.Text) Then
        If txt_fuenteMedidas.Text > 5 Then
            str_valida = str_valida & vbCrLf & "La fuente de Medida debe ser letra(a-z) o un numero menor a 6."
            txt_fuenteMedidas.SetFocus
        End If
    Else
        If Len(txt_fuenteMedidas.Text) > 1 Then
            str_valida = str_valida & vbCrLf & "Solo se acepta una letra para fuente de Medida"
            txt_fuenteMedidas.SetFocus
        End If
    End If
    If IsNumeric(txt_fuentePrecio.Text) Then
        If txt_fuentePrecio.Text > 5 Then
            str_valida = str_valida & vbCrLf & "La fuente de Precio debe ser letra(a-z) o un numero menor a 6."
            txt_fuentePrecio.SetFocus
        End If
    Else
        If Len(txt_fuentePrecio.Text) > 1 Then
            str_valida = str_valida & vbCrLf & "Solo se acepta una letra para fuente de Precio"
            txt_fuentePrecio.SetFocus
        End If
    End If
    If IsNumeric(txt_fuenteUbicacion.Text) Then
        If txt_fuenteUbicacion.Text > 5 Then
            str_valida = str_valida & vbCrLf & "La fuente de Ubicacion debe ser letra(a-z) o un numero menor a 6."
            txt_fuenteUbicacion.SetFocus
        End If
    Else
        If Len(txt_fuenteUbicacion.Text) > 1 Then
            str_valida = str_valida & vbCrLf & "Solo se acepta una letra para fuente de Ubicacion."
            txt_fuenteUbicacion.SetFocus
        End If
    End If
    
    If IsNumeric(txt_ancho.Text) Then
        If txt_ancho.Text <= 0 Then
            str_valida = str_valida & vbCrLf & "La cantidad  debe ser mayor a ''Cero'' para el ancho."
            txt_ancho.SetFocus
        End If
    Else
        str_valida = str_valida & vbCrLf & "Solo se acepta numero para el ancho."
        txt_ancho.SetFocus
    End If
    
    If IsNumeric(txt_largo.Text) Then
        If txt_largo.Text <= 0 Then
            str_valida = str_valida & vbCrLf & "La cantidad  debe ser mayor a ''Cero'' para el largo."
            txt_largo.SetFocus
        End If
    Else
        str_valida = str_valida & vbCrLf & "Solo se acepta numero para el largo."
        txt_largo.SetFocus
    End If
    
    If Not IsNumeric(txt_lineaNegra.Text) Then
        str_valida = str_valida & vbCrLf & "Solo se acepta numero para la linea negra."
        txt_lineaNegra.SetFocus
    End If
    If txt_descripcion.Text = "" Then
        If txt_codigo.Text = "" Then
            str_valida = str_valida & vbCrLf & "Favor de capturar el nombre del Tipo de Etiquete."
            txt_descripcion.Locked = False
            txt_descripcion.SetFocus
        End If
    End If
    
End Function

Private Sub cmd_guardar_Click()
    If Conectar_BD(cnnSID, "compucaja", "srvtdacantia") Then
        Dim cmdGuarda As New ADODB.Command
        cmdGuarda.ActiveConnection = cnnSID
        cmdGuarda.CommandText = "PC_GuardaConfiguracionEtiquetas"
        cmdGuarda.CommandType = adCmdStoredProc
        cmdGuarda("@inte_eti_etiqueta_id").Value = txt_codigo.Text
        cmdGuarda("@vcha_eti_etiqueta").Value = txt_descripcion.Text
        cmdGuarda("@floa_eti_ancho").Value = txt_ancho.Text
        cmdGuarda("@floa_eti_largo").Value = txt_largo.Text
        cmdGuarda("@vcha_eti_fuenteEncabezado1").Value = txt_FuenteEnc1.Text
        cmdGuarda("@vcha_eti_fuenteEncabezado2").Value = txt_fuenteEnc2.Text
        cmdGuarda("@vcha_eti_fuentePrecio").Value = txt_fuentePrecio.Text
        cmdGuarda("@vcha_eti_fuenteCodigo1").Value = txt_fuenteCodigo.Text
        cmdGuarda("@vcha_eti_fuenteCodigo2").Value = txt_fuentecolor.Text
        cmdGuarda("@vcha_eti_fuenteMedida").Value = txt_fuenteMedidas.Text
        cmdGuarda("@vcha_eti_fuenteUbicacion").Value = txt_fuenteUbicacion.Text
        cmdGuarda("@vcha_eti_fuenteSignoPrecio").Value = txt_fuenteSignoPrecio.Text
        
        cmdGuarda("@bint_eti_posi_x_Encabezado1").Value = txt_Posi_x_Enc1.Text
        cmdGuarda("@bint_eti_posi_x_Encabezado2").Value = txt_Posi_x_Enc2.Text
        cmdGuarda("@bint_eti_posi_x_Presio").Value = txt_Posi_x_Precio.Text
        cmdGuarda("@bint_eti_posi_x_Codigo1").Value = txt_Posi_x_Codigo.Text
        cmdGuarda("@bint_eti_posi_x_Codigo2").Value = txt_Posi_x_Color.Text
        cmdGuarda("@bint_eti_posi_x_Medida").Value = txt_Posi_x_Medidas.Text
        cmdGuarda("@bint_eti_posi_x_Ubicacion").Value = txt_Posi_x_Loc.Text
        cmdGuarda("@bint_eti_posi_x_SignoPrecio").Value = txt_Posi_x_SignoPrecio.Text
        
        If cbo_rotacion.Text = "Sin Rotacion" Then
            cmdGuarda("@bint_eti_posicion").Value = 0
        Else
            cmdGuarda("@bint_eti_posicion").Value = 1
        End If
        cmdGuarda("@bint_eti_posi_y_Encabezado1").Value = txt_Posi_y_Enc1.Text
        cmdGuarda("@bint_eti_posi_y_Encabezado2").Value = txt_Posi_y_Enc2.Text
        cmdGuarda("@bint_eti_posi_y_Presio").Value = txt_Posi_y_Precio.Text
        cmdGuarda("@bint_eti_posi_y_Codigo1").Value = txt_Posi_y_Codigo.Text
        cmdGuarda("@bint_eti_posi_y_Codigo2").Value = txt_Posi_y_Color.Text
        cmdGuarda("@bint_eti_posi_y_Ubicacion").Value = txt_Posi_y_Loc.Text
        cmdGuarda("@bint_eti_posi_y_Medida").Value = txt_Posi_y_Medidas.Text
        cmdGuarda("@bint_eti_posi_y_SignoPrecio").Value = txt_Posi_y_SignoPrecio.Text
        
        
        cmdGuarda("@bint_eti_CaracteresEncabezado1").Value = txt_CaracteresEnc1.Text
        cmdGuarda("@int_eti_CaracteresEncabezado2").Value = txt_CaracteresEnc2.Text
        cmdGuarda("@int_eti_CaracteresPresio").Value = txt_CaracteresPrecio.Text
        cmdGuarda("@int_eti_CaracteresCodigo1").Value = txt_CaracteresCodigo.Text
        cmdGuarda("@int_eti_CaracteresCodigo2").Value = txt_CaracteresColor.Text
        cmdGuarda("@int_eti_CaracteresMedida").Value = txt_CaracteresMedidas.Text
        cmdGuarda("@bint_eti_CaracteresUbicacion").Value = txt_CaracteresLoc.Text
        
        
        cmdGuarda("@int_eti_TamLineaEncabezado1").Value = txt_tamLineaEnc1.Text
        cmdGuarda("@int_eti_TamLineaEncabezado2").Value = txt_tamLineaEnc2.Text
        cmdGuarda("@int_eti_TamLineaPresio").Value = txt_tamLineaPrecio.Text
        cmdGuarda("@int_eti_TamLineaCodigo1").Value = txt_tamLineaCodigo.Text
        cmdGuarda("@int_eti_TamLineaCodigo2").Value = txt_tamLineaColor.Text
        cmdGuarda("@bint_eti_TamLineaMedida").Value = txt_tamLineaMedia.Text
        cmdGuarda("@int_eti_TamLineaUbicacion").Value = txt_tamLineaLoc.Text
        cmdGuarda("@floa_eti_TamLineaNegra").Value = txt_lineaNegra.Text
        If txt_lineaNegra.Text > 0 Then
            cmdGuarda("@vcha_eti_lineaNegra").Value = "B"
        Else
            cmdGuarda("@vcha_eti_lineaNegra").Value = ""
        End If
        cmdGuarda("@vcha_eti_campoEncabezado1").Value = txt_CampoEncabezado1.Text
        cmdGuarda("@vcha_eti_campoEncabezado2").Value = txt_CampoEncabezado2.Text
        
        cmdGuarda.execute
        MsgBox cmdGuarda("@var_mensaje").Value, vbInformation, "SID"
    Else
        MsgBox "No se puede conectar a la base de datos", vbCritical, "SID"
    End If
    cnnSID.Close
End Sub
Private Sub pro_limpiaTodo()
    txt_ancho.Text = ""
    txt_CaracteresCodigo.Text = ""
    txt_CaracteresColor.Text = ""
    txt_CaracteresEnc1.Text = ""
    txt_CaracteresEnc2.Text = ""
    txt_CaracteresLoc.Text = ""
    txt_CaracteresMedidas.Text = ""
    txt_CaracteresPrecio.Text = ""
    txt_codigo.Text = ""
    txt_descripcion.Text = ""
    txt_descripcion.Locked = False
    txt_fuenteCodigo.Text = ""
    txt_fuentecolor.Text = ""
    txt_FuenteEnc1.Text = ""
    txt_fuenteEnc2.Text = ""
    txt_fuenteMedidas.Text = ""
    txt_fuentePrecio.Text = ""
    txt_fuenteUbicacion.Text = ""
    txt_largo.Text = ""
    txt_lineaNegra.Text = ""
    txt_Posi_x_Codigo.Text = ""
    txt_Posi_x_Color.Text = ""
    txt_Posi_x_Enc1.Text = ""
    txt_Posi_x_Enc2.Text = ""
    txt_Posi_x_Loc.Text = ""
    txt_Posi_x_Medidas.Text = ""
    txt_Posi_x_Precio.Text = ""
    txt_Posi_y_Codigo.Text = ""
    txt_Posi_y_Color.Text = ""
    txt_Posi_y_Enc1.Text = ""
    txt_Posi_y_Enc2.Text = ""
    txt_Posi_y_Loc.Text = ""
    txt_Posi_y_Medidas.Text = ""
    txt_Posi_y_Precio.Text = ""
    txt_tamLineaCodigo.Text = ""
    txt_tamLineaColor.Text = ""
    txt_tamLineaEnc1.Text = ""
    txt_tamLineaEnc2.Text = ""
    txt_tamLineaLoc.Text = ""
    txt_tamLineaMedia.Text = ""
    txt_tamLineaPrecio.Text = ""
    cbo_rotacion.Text = "Sin Rotacion"
    
    
    
End Sub



Private Sub cmd_nuevo_Click()
    pro_limpiaTodo
End Sub

Private Sub txt_codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Dim frmBuscaFuente As New frmBucarArticulos
        
        With frmBuscaFuente
            .catalogo = " de Etiquetas de Precios"
           
            .sentenciaSql = "SELECT inte_eti_etiqueta_id id , vcha_eti_etiqueta  nombre " & _
                    "FROM tb_etiquetasConfiguracion WITH (NOLOCK) " & _
                    "WHERE (inte_eti_etiqueta_id  LIKE '%XXXXXX%' " & _
                    "OR vcha_eti_etiqueta LIKE '%XXXXXX%') "
        
            .comodin = "XXXXXX"
        
            .Show vbModal
            'Revisar si se seleccionó alguna clase de artículo.
            If .valorSeleccionado1 <> "" Then
                Call pro_cargaInformacion(.valorSeleccionado1)
                Call pro_HabilitaComponentes
            End If
        
        
        End With
    End If
End Sub

Private Sub pro_HabilitaComponentes()
    txt_ancho.Locked = False
    txt_ancho.Enabled = True
    txt_CaracteresCodigo.Locked = False
    txt_CaracteresColor.Locked = False
    txt_CaracteresEnc1.Locked = False
    txt_CaracteresEnc2.Locked = False
    txt_CaracteresLoc.Locked = False
    txt_CaracteresMedidas.Locked = False
    txt_CaracteresPrecio.Locked = False
    txt_descripcion.Locked = False
    txt_codigo.Locked = True
    txt_fuenteCodigo.Locked = False
    txt_fuentecolor.Locked = False
    txt_FuenteEnc1.Locked = False
    txt_fuenteEnc2.Locked = False
    txt_fuenteMedidas.Locked = False
    txt_fuentePrecio.Locked = False
    txt_fuenteUbicacion.Locked = False
    txt_fuenteSignoPrecio.Locked = False
    txt_largo.Locked = False
    txt_largo.Enabled = True
    txt_lineaNegra.Locked = False
    txt_lineaNegra.Enabled = True
    txt_Posi_x_Codigo.Locked = False
    txt_Posi_x_Color.Locked = False
    txt_Posi_x_Enc1.Locked = False
    txt_Posi_x_Enc2.Locked = False
    txt_Posi_x_Loc.Locked = False
    txt_Posi_x_Medidas.Locked = False
    txt_Posi_x_Precio.Locked = False
    txt_Posi_x_SignoPrecio.Locked = False
    txt_Posi_y_Codigo.Locked = False
    txt_Posi_y_Color.Locked = False
    txt_Posi_y_Enc1.Locked = False
    txt_Posi_y_Enc2.Locked = False
    txt_Posi_y_Loc.Locked = False
    txt_Posi_y_Medidas.Locked = False
    txt_Posi_y_Precio.Locked = False
    txt_Posi_y_SignoPrecio.Locked = False
    txt_tamLineaCodigo.Locked = False
    txt_tamLineaColor.Locked = False
    txt_tamLineaEnc1.Locked = False
    txt_tamLineaEnc2.Locked = False
    txt_tamLineaLoc.Locked = False
    txt_tamLineaMedia.Locked = False
    txt_tamLineaPrecio.Locked = False
    txt_CampoEncabezado1.Locked = False
    txt_CampoEncabezado2.Locked = False
    cbo_rotacion.Enabled = True
    
End Sub

Private Sub pro_cargaInformacion(strEtiqueta_ID As String)
    If Conectar_BD(cnnSID, "compucaja", "srvtdacantia") Then
        Dim rsEtiquetas As New ADODB.recordSet
        rsEtiquetas.Open "Select * " & _
                        "from tb_etiquetasConfiguracion with(nolock) " & _
                        "where inte_eti_etiqueta_id = " & strEtiqueta_ID, _
                    cnnSID, _
                    adOpenDynamic, _
                    adLockOptimistic
        If rsEtiquetas.RecordCount > 0 Then
            txt_descripcion.Locked = True
            txt_codigo.Text = strEtiqueta_ID
            txt_descripcion.Text = rsEtiquetas("vcha_eti_etiqueta").Value
            txt_lineaNegra.Text = rsEtiquetas("floa_eti_TamLineaNegra").Value
            
            
            txt_ancho.Text = rsEtiquetas("floa_eti_ancho").Value
            txt_largo.Text = rsEtiquetas("floa_eti_largo").Value
            
            
            txt_FuenteEnc1.Text = rsEtiquetas("vcha_eti_fuenteEncabezado1").Value
            txt_fuenteEnc2.Text = rsEtiquetas("vcha_eti_fuenteEncabezado2").Value
            txt_fuentePrecio.Text = rsEtiquetas("vcha_eti_fuentePrecio").Value
            txt_fuenteCodigo.Text = rsEtiquetas("vcha_eti_fuenteCodigo1").Value
            txt_fuentecolor.Text = rsEtiquetas("vcha_eti_fuenteCodigo2").Value
            txt_fuenteMedidas.Text = rsEtiquetas("vcha_eti_fuenteMedida").Value
            txt_fuenteUbicacion.Text = rsEtiquetas("vcha_eti_fuenteUbicacion").Value
            txt_fuenteSignoPrecio.Text = IIf(IsNull(rsEtiquetas("vcha_eti_fuenteSignoPrecio").Value), "", rsEtiquetas("vcha_eti_fuenteSignoPrecio").Value)
            
            
            txt_Posi_x_Enc1.Text = rsEtiquetas("bint_eti_posi_x_Encabezado1").Value
            txt_Posi_x_Enc2.Text = rsEtiquetas("bint_eti_posi_x_Encabezado2").Value
            txt_Posi_x_Precio.Text = rsEtiquetas("bint_eti_posi_x_Presio").Value
            txt_Posi_x_Codigo.Text = rsEtiquetas("bint_eti_posi_x_Codigo1").Value
            txt_Posi_x_Color.Text = rsEtiquetas("bint_eti_posi_x_Codigo2").Value
            txt_Posi_x_Medidas.Text = rsEtiquetas("bint_eti_posi_x_Medida").Value
            txt_Posi_x_Loc.Text = rsEtiquetas("bint_eti_posi_x_Ubicacion").Value
            txt_Posi_x_SignoPrecio.Text = IIf(IsNull(rsEtiquetas("bint_eti_posi_x_SignoPrecio").Value), "", rsEtiquetas("bint_eti_posi_x_SignoPrecio").Value)
            If rsEtiquetas("bint_eti_posicion").Value = 0 Then
                cbo_rotacion.Text = "Sin Rotacion"
            Else
                cbo_rotacion.Text = "90 Grados"
            End If
            
            txt_Posi_y_Enc1.Text = rsEtiquetas("bint_eti_posi_y_Encabezado1").Value
            txt_Posi_y_Enc2.Text = rsEtiquetas("bint_eti_posi_y_Encabezado2").Value
            txt_Posi_y_Precio.Text = rsEtiquetas("bint_eti_posi_y_Presio").Value
            txt_Posi_y_Codigo.Text = rsEtiquetas("bint_eti_posi_y_Codigo1").Value
            txt_Posi_y_Color.Text = rsEtiquetas("bint_eti_posi_y_Codigo2").Value
            txt_Posi_y_Loc.Text = rsEtiquetas("bint_eti_posi_y_Ubicacion").Value
            txt_Posi_y_Medidas.Text = rsEtiquetas("bint_eti_posi_y_Medida").Value
            txt_Posi_y_SignoPrecio.Text = IIf(IsNull(rsEtiquetas("bint_eti_posi_y_SignoPrecio").Value), "", rsEtiquetas("bint_eti_posi_y_SignoPrecio").Value)
            
            
            txt_CaracteresEnc1.Text = rsEtiquetas("bint_eti_CaracteresEncabezado1").Value
            txt_CaracteresEnc2.Text = rsEtiquetas("int_eti_CaracteresEncabezado2").Value
            txt_CaracteresPrecio.Text = rsEtiquetas("int_eti_CaracteresPresio").Value
            txt_CaracteresCodigo.Text = rsEtiquetas("int_eti_CaracteresCodigo1").Value
            txt_CaracteresColor.Text = rsEtiquetas("int_eti_CaracteresCodigo2").Value
            txt_CaracteresMedidas.Text = rsEtiquetas("int_eti_CaracteresMedida").Value
            txt_CaracteresLoc.Text = rsEtiquetas("bint_eti_CaracteresUbicacion").Value
            
            txt_tamLineaEnc1.Text = rsEtiquetas("int_eti_TamLineaEncabezado1").Value
            txt_tamLineaEnc2.Text = rsEtiquetas("int_eti_TamLineaEncabezado2").Value
            txt_tamLineaPrecio.Text = rsEtiquetas("int_eti_TamLineaPresio").Value
            txt_tamLineaCodigo.Text = rsEtiquetas("int_eti_TamLineaCodigo1").Value
            txt_tamLineaColor.Text = rsEtiquetas("int_eti_TamLineaCodigo2").Value
            txt_tamLineaMedia.Text = rsEtiquetas("bint_eti_TamLineaMedida").Value
            txt_tamLineaLoc.Text = rsEtiquetas("int_eti_TamLineaUbicacion").Value
            txt_tamLineaEnc1.Text = rsEtiquetas("int_eti_TamLineaEncabezado1").Value
            
            txt_CampoEncabezado1.Text = rsEtiquetas("vcha_eti_campoEncabezado1").Value
            txt_CampoEncabezado2.Text = rsEtiquetas("vcha_eti_campoEncabezado2").Value
            
        Else
            MsgBox "No se encontró informacion ", vbCritical, "SID"
        End If
        rsEtiquetas.Close
        cnnSID.Close
    Else
        MsgBox "No se puede conectar a la base de datos", vbCritical, "SID"
    End If
    
End Sub

Private Function Conectar_BD(ByRef cnnCBD As ADODB.Connection, ByVal bd As String, ByVal servidor As String) As Boolean
    'Variables de bloque
    Dim strConnection_String As String
    
On Error GoTo Error_Conectar_BDS
    Conectar_BD = True
    'Establecer connection strings para realizar las conexiones a las bases de
    'datos
    If cnnSID.State = 1 Then
        cnnSID.Close
    End If
    
    
    strConnection_String_SID = "Provider=SQLOLEDB.1;Password=compucaja" & _
                                ";Persist Security Info=True;User ID=sa" & _
                                ";Initial Catalog=" & UCase(bd) & ";Data Source=" & UCase(servidor)
    
    'Configurar objetos Connection
    'cnnCBD.CursorLocation = adUseClient
    If cnnCBD.State = 1 Then
        cnnCBD.Close
    End If
    cnnCBD.ConnectionString = strConnection_String_SID
    cnnCBD.CommandTimeout = 60
    cnnCBD.CursorLocation = adUseClient
    
    'Abrir conexiones a las bases de datos
    cnnCBD.Open
    Exit Function
Error_Conectar_BDS:
    Conectar_BD = False
    MsgBox Err.Description, vbCritical, "SID"
End Function

