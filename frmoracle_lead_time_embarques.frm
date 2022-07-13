VERSION 5.00
Begin VB.Form frmoracle_lead_time_embarques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tiempo"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_fin 
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   18
      Top             =   1440
      Width           =   8415
   End
   Begin VB.TextBox txt_unidad 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   600
      Width           =   5415
   End
   Begin VB.TextBox txt_inicio 
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txt_fin_lectores 
      Height          =   405
      Left            =   5400
      TabIndex        =   13
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txt_inicio_lectores 
      Height          =   405
      Left            =   2280
      TabIndex        =   12
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txt_fin_facturas 
      Height          =   405
      Left            =   5400
      TabIndex        =   7
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txt_inicio_facturas 
      Height          =   405
      Left            =   2280
      TabIndex        =   6
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txt_fin_notas 
      Height          =   405
      Left            =   5400
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txt_inicio_notas 
      Height          =   405
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txt_fin_aduanales 
      Height          =   405
      Left            =   5400
      TabIndex        =   3
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txt_inicio_aduanales 
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txt_embarque 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Fin:"
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
      Left            =   3960
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Unidad:"
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
      TabIndex        =   16
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label6 
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
      TabIndex        =   14
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Facturas:"
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
      TabIndex        =   11
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Notas de envio:"
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
      TabIndex        =   10
      Top             =   2640
      Width           =   1875
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Aduanales:"
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
      TabIndex        =   9
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lectores:"
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
      TabIndex        =   8
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Embarque:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1320
   End
End
Attribute VB_Name = "frmoracle_lead_time_embarques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_Change()
    Me.txt_fin_aduanales = ""
    Me.txt_fin_facturas = ""
    Me.txt_fin_lectores = ""
    Me.txt_fin_notas = ""
    Me.txt_inicio_aduanales = ""
    Me.txt_inicio_facturas = ""
    Me.txt_inicio_lectores = ""
    Me.txt_inicio_notas = ""
    Me.txt_unidad = ""
    Me.txt_inicio = ""
    Me.txt_fin = ""
    
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If IsNumeric(Me.txt_embarque) Then
      var_cadena = "select * from xxvia_Tb_Encabezado_embarques where embarque = ?"
      With comandoORA
           .ActiveConnection = cnnoracle_4
           .CommandType = adCmdText
           .CommandText = var_cadena
           Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CStr(Me.txt_embarque))
           .Parameters.Append parametro
      End With
      Set rsaux = comandoORA.execute
      Set comandoORA = Nothing
      Set parametro = Nothing
      If Not rsaux.EOF Then
         Me.txt_unidad = IIf(IsNull(rsaux!VEHICULO), "", rsaux!VEHICULO)
         Me.txt_inicio = IIf(IsNull(rsaux!FECHA_INiCIO), "", rsaux!FECHA_INiCIO)
         Me.txt_fin = IIf(IsNull(rsaux!FECHA_FIN), "", rsaux!FECHA_FIN)
         
         If rs.State = 1 Then
            rs.Close
         End If
         var_cadena = "select min(hora_inicio), max(hora_final) from TB_ORACLE_TIEMPO_POR_LOTE where pedido in (select pedido from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = " + Me.txt_embarque + ")"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_inicio_lectores = IIf(IsNull(rs(0).Value), "", rs(0).Value)
            Me.txt_fin_lectores = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         Else
            Me.txt_inicio_lectores = ""
            Me.txt_fin_lectores = ""
         End If
         If rs.State = 1 Then
            rs.Close
         End If
         var_cadena = "select min(hora_inicio), max(hora_fin) from TB_ORACLE_TIEMPO_PEDIDO_ADUANAS where pedido in (select pedido from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = " + Me.txt_embarque + ")"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_inicio_aduanales = IIf(IsNull(rs(0).Value), "", rs(0).Value)
            Me.txt_fin_aduanales = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         Else
            Me.txt_inicio_aduanales = ""
            Me.txt_fin_aduanales = ""
         End If
         If rs.State = 1 Then
            rs.Close
         End If
         var_cadena = "SELECT min(fecha), max(fecha) FROM TB_ORACLE_TIEMPO_IMPRESION_DOCUMENTOS where pedido in (select pedido from TB_ORACLE_PEDIDOS_ASIGNADOS_EMBARQUES where embarque = " + Me.txt_embarque + ")"
         rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            Me.txt_inicio_notas = IIf(IsNull(rs(0).Value), "", rs(0).Value)
            Me.txt_fin_notas = IIf(IsNull(rs(0).Value), "", rs(0).Value)
         Else
            Me.txt_inicio_notas = ""
            Me.txt_fin_notas = ""
         End If
         If rs.State = 1 Then
            rs.Close
         End If
         rs.Open "alter session set nls_date_format = 'DD-MON-YYYY HH24:MI:SS'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         var_cadena = "select min(creation_Date), max(creation_Date) from ra_customer_Trx_all where ct_reference in (select distinct to_char(source_header_number) from xxvia_Tb_Salidas_cajas where inte_Emb_Embarque = ?)"
         With comandoORA
              .ActiveConnection = cnnoracle_4
              .CommandType = adCmdText
              .CommandText = var_cadena
              Set parametro = .CreateParameter(, adNumeric, adParamInput, 100, CStr(Me.txt_embarque))
              .Parameters.Append parametro
         End With
         Set rs = comandoORA.execute
         Set comandoORA = Nothing
         Set parametro = Nothing
         If Not rs.EOF Then
            Me.txt_inicio_facturas = IIf(IsNull(rs(0).Value), "", rs(0).Value)
            Me.txt_fin_facturas = IIf(IsNull(rs(1).Value), "", rs(1).Value)
         Else
            Me.txt_inicio_facturas = ""
            Me.txt_fin_facturas = ""
         End If
      Else
         MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
      End If
   End If
End Sub
