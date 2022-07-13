VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmoracle_volumen_embarque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Volumen por embarque"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   2880
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   3960
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame frm_embarque 
      Height          =   735
      Left            =   4800
      TabIndex        =   19
      Top             =   960
      Width           =   2175
      Begin VB.TextBox txt_embarque_buscar 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame frm_lista 
      Height          =   2400
      Left            =   1320
      TabIndex        =   16
      Top             =   240
      Width           =   7365
      Begin MSComctlLib.ListView lv_lista 
         Height          =   1830
         Left            =   0
         TabIndex        =   17
         Top             =   480
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   3228
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clave"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7585
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Volumen"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_lista 
         BackColor       =   &H000000FF&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   7290
      End
   End
   Begin VB.TextBox txt_embarques 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   600
      Width           =   6375
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Picture         =   "frmoracle_volumen_embarque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Actualizar "
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txt_porcentaje_carga 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txt_volumen_carga 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmd_siguiente 
      Caption         =   "siguiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd_anterior 
      Caption         =   "anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_volumen_unidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txt_nombre_unidad 
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
      Left            =   3285
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   5535
   End
   Begin VB.TextBox txt_clave_unidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txt_embarque 
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
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txt_porcentaje_carga_lectores 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txt_volumen_carga_lectores 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Vol. Lectores:"
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
      Left            =   3360
      TabIndex        =   25
      Top             =   2340
      Width           =   1680
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "% Lectores:"
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
      Left            =   6360
      TabIndex        =   24
      Top             =   2340
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Embarques:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "% Aduana:"
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
      Left            =   6480
      TabIndex        =   12
      Top             =   3420
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vol. Aduana:"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   3420
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Volumen unidad:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2340
      Width           =   2025
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmoracle_volumen_embarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comandoORA As New ADODB.Command
Dim parametro As ADODB.Parameter

Private Sub Command4_Click()
   If IsNumeric(Me.txt_embarque) Then
      rs.Open "select * from TB_ORACLE_GRUPOS_EMBARQUES where grupo = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
      var_Cadena_embarques = ""
      If Not rs.EOF Then
         var_transporte = IIf(IsNull(rs!unidad), "", rs!unidad)
         While Not rs.EOF
               If var_Cadena_embarques = "" Then
                  var_Cadena_embarques = rs!Embarque
               Else
                  var_Cadena_embarques = var_Cadena_embarques + "," + rs!Embarque
               End If
               rs.MoveNext
         Wend
         rs.Close
         Me.txt_embarques = var_Cadena_embarques
          
         strconsulta = "SELECT * FROM XXVIA_tB_ENCABEZADO_EMBARQUES WHERE embarque in (" + var_Cadena_embarques + ") "
         rs.Open strconsulta, cnnoracle_4, adOpenDynamic, adLockOptimistic
         'With comandoORA
         '     .ActiveConnection = cnnoracle_4
         '     .CommandType = adCmdText
         '     .CommandText = strconsulta
         '     Set parametro = .CreateParameter(, adVarChar, adParamInput, 1000, var_Cadena_embarques)
         '     .Parameters.Append parametro
         'End With
         'MsgBox var_Cadena_embarques
         'Set rs = comandoORA.execute
         'Set comandoORA = Nothing
         'Set parametro = Nothing
         
         rsaux1.Open "SELECT SUM(TB_ORACLE_EMPAQUES.VOLUMEN) FROM TB_ORACLE_CAJAS_ADUANA A, TB_ORACLE_EMPAQUES WHERE EMBARQUE in (" + var_Cadena_embarques + ") AND TIPO_EMPAQUE = TB_ORACLE_EMPAQUES.EMPAQUE AND ESTATUS in ('S','L')", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_volumen_carga = Round(IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value), 2)
         Else
            Me.txt_volumen_carga = 0
         End If
         rsaux1.Close
         
        
         
         rsaux1.Open "SELECT SUM(TB_ORACLE_EMPAQUES.VOLUMEN) FROM TB_ORACLE_CAJAS_ADUANA A, TB_ORACLE_EMPAQUES WHERE EMBARQUE in (" + var_Cadena_embarques + ") AND TIPO_EMPAQUE = TB_ORACLE_EMPAQUES.EMPAQUE", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux1.EOF Then
            Me.txt_volumen_carga_lectores = Round(IIf(IsNull(rsaux1(0).Value), 0, rsaux1(0).Value), 2)
         Else
            Me.txt_volumen_carga_lectores = 0
         End If
         rsaux1.Close
         
         
         
         
         If Not rs.EOF Then
            rsaux.Open "select * from tb_oracle_transportes where clave = '" + var_transporte + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               Me.txt_clave_unidad = IIf(IsNull(rsaux!CLAVE), "", rsaux!CLAVE)
               Me.txt_nombre_unidad = IIf(IsNull(rsaux!nombre), "", rsaux!nombre)
               Me.txt_volumen_unidad = Round(IIf(IsNull(rsaux!VOLUMEN), 0, rsaux!VOLUMEN), 2)
               
               If CDbl(Me.txt_volumen_unidad) > 0 Then
                  var_porcentaje = (CDbl(Me.txt_volumen_carga) * 100) / CDbl(Me.txt_volumen_unidad)
                  Me.txt_porcentaje_carga = Round(var_porcentaje, 2)
               Else
                  var_porcentaje = 0
                  Me.txt_porcentaje_carga = Round(var_porcentaje, 2)
               End If
               
               
               
               If var_porcentaje = 0 Then
                  Me.ProgressBar1.Value = 0
               End If
               
               If var_porcentaje > 0 And var_porcentaje < 5.01 Then
                  Me.ProgressBar1.Value = 5
               End If
               
               If var_porcentaje > 5 And var_porcentaje < 10.01 Then
                  Me.ProgressBar1.Value = 10
               End If
               
               If var_porcentaje > 10 And var_porcentaje < 15.01 Then
                  Me.ProgressBar1.Value = 15
               End If
               
               If var_porcentaje > 15 And var_porcentaje < 20.01 Then
                  Me.ProgressBar1.Value = 20
               End If
               
               If var_porcentaje > 20 And var_porcentaje < 25.01 Then
                  Me.ProgressBar1.Value = 25
               End If
               
               
               If var_porcentaje > 25 And var_porcentaje < 30.01 Then
                  Me.ProgressBar1.Value = 30
               End If
               
               
               If var_porcentaje > 30 And var_porcentaje < 35.01 Then
                  Me.ProgressBar1.Value = 35
               End If
               
               
               If var_porcentaje > 35 And var_porcentaje < 40.01 Then
                  Me.ProgressBar1.Value = 40
               End If
               
               If var_porcentaje > 40 And var_porcentaje < 45.01 Then
                  Me.ProgressBar1.Value = 45
               End If
               
               If var_porcentaje > 45 And var_porcentaje < 50.01 Then
                  Me.ProgressBar1.Value = 50
               End If
               
               If var_porcentaje > 50 And var_porcentaje < 55.01 Then
                  Me.ProgressBar1.Value = 55
               End If
               
               If var_porcentaje > 55 And var_porcentaje < 60.01 Then
                  Me.ProgressBar1.Value = 60
               End If
               
               
               If var_porcentaje > 60 And var_porcentaje < 65.01 Then
                  Me.ProgressBar1.Value = 65
               End If
               
               If var_porcentaje > 65 And var_porcentaje < 70.01 Then
                  Me.ProgressBar1.Value = 70
               End If
               
               If var_porcentaje > 70 And var_porcentaje < 75.01 Then
                  Me.ProgressBar1.Value = 75
               End If
               
               
               If var_porcentaje > 75 And var_porcentaje < 80.01 Then
                  Me.ProgressBar1.Value = 80
               End If
               
               If var_porcentaje > 80 And var_porcentaje < 85.01 Then
                  Me.ProgressBar1.Value = 85
               End If
               
               If var_porcentaje > 85 And var_porcentaje < 90.01 Then
                  Me.ProgressBar1.Value = 90
               End If
               
               
               If var_porcentaje > 90 And var_porcentaje < 95.01 Then
                  Me.ProgressBar1.Value = 95
               End If
               
               If var_porcentaje > 95 Then
                  Me.ProgressBar1.Value = 100
               End If
'------------------
               If CDbl(Me.txt_volumen_unidad) > 0 Then
                  var_porcentaje = (CDbl(Me.txt_volumen_carga_lectores) * 100) / CDbl(Me.txt_volumen_unidad)
                  Me.txt_porcentaje_carga_lectores = Round(var_porcentaje, 2)
               Else
                  var_porcentaje = 0
                  Me.txt_porcentaje_carga_lectores = Round(var_porcentaje, 2)
               End If
               
               
               If var_porcentaje = 0 Then
                  Me.ProgressBar2.Value = 0
               End If
               
               If var_porcentaje > 0 And var_porcentaje < 5.01 Then
                  Me.ProgressBar2.Value = 5
               End If
               
               If var_porcentaje > 5 And var_porcentaje < 10.01 Then
                  Me.ProgressBar2.Value = 10
               End If
               
               If var_porcentaje > 10 And var_porcentaje < 15.01 Then
                  Me.ProgressBar2.Value = 15
               End If
               
               If var_porcentaje > 15 And var_porcentaje < 20.01 Then
                  Me.ProgressBar2.Value = 20
               End If
               
               If var_porcentaje > 20 And var_porcentaje < 25.01 Then
                  Me.ProgressBar2.Value = 25
               End If
               
               
               If var_porcentaje > 25 And var_porcentaje < 30.01 Then
                  Me.ProgressBar2.Value = 30
               End If
               
               
               If var_porcentaje > 30 And var_porcentaje < 35.01 Then
                  Me.ProgressBar2.Value = 35
               End If
               
               
               If var_porcentaje > 35 And var_porcentaje < 40.01 Then
                  Me.ProgressBar2.Value = 40
               End If
               
               If var_porcentaje > 40 And var_porcentaje < 45.01 Then
                  Me.ProgressBar2.Value = 45
               End If
               
               If var_porcentaje > 45 And var_porcentaje < 50.01 Then
                  Me.ProgressBar2.Value = 50
               End If
               
               If var_porcentaje > 50 And var_porcentaje < 55.01 Then
                  Me.ProgressBar2.Value = 55
               End If
               
               If var_porcentaje > 55 And var_porcentaje < 60.01 Then
                  Me.ProgressBar2.Value = 60
               End If
               
               
               If var_porcentaje > 60 And var_porcentaje < 65.01 Then
                  Me.ProgressBar2.Value = 65
               End If
               
               If var_porcentaje > 65 And var_porcentaje < 70.01 Then
                  Me.ProgressBar2.Value = 70
               End If
               
               If var_porcentaje > 70 And var_porcentaje < 75.01 Then
                  Me.ProgressBar2.Value = 75
               End If
               
               
               If var_porcentaje > 75 And var_porcentaje < 80.01 Then
                  Me.ProgressBar2.Value = 80
               End If
               
               If var_porcentaje > 80 And var_porcentaje < 85.01 Then
                  Me.ProgressBar2.Value = 85
               End If
               
               If var_porcentaje > 85 And var_porcentaje < 90.01 Then
                  Me.ProgressBar2.Value = 90
               End If
               
               
               If var_porcentaje > 90 And var_porcentaje < 95.01 Then
                  Me.ProgressBar2.Value = 95
               End If
               
               If var_porcentaje > 95 Then
                  Me.ProgressBar2.Value = 100
               End If
               
            Else
               MsgBox "El grupo de embarques no tienen una unidad seleccionada", vbOKOnly, "ATENCION"
               Me.txt_clave_unidad = ""
               Me.txt_nombre_unidad = ""
               Me.txt_volumen_unidad = 0
               Me.txt_porcentaje_carga = 0
               Me.txt_porcentaje_carga_lectores = 0
               
               Me.txt_clave_unidad.SetFocus
            End If
            rsaux.Close
         Else
            MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      Else
         rs.Close
         MsgBox "El grupo de embarques no existe"
      End If
   Else
      MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Form_Load()
   Me.frm_embarque.Visible = False
   Me.frm_lista.Visible = False
   Top = 1000
   Left = 1500
   'Me.img_5.Visible = True
   'Me.img_10.Visible = False
   'Me.img_15.Visible = False
   'Me.img_20.Visible = False
   'Me.img_25.Visible = False
   'Me.img_30.Visible = False
   'Me.img_35.Visible = False
   'Me.img_40.Visible = False
   'Me.img_45.Visible = False
   'Me.img_50.Visible = False
   'Me.img_55.Visible = False
   'Me.img_60.Visible = False
   'Me.img_65.Visible = False
   'Me.img_70.Visible = False
   'Me.img_75.Visible = False
   'Me.img_80.Visible = False
   'Me.img_85.Visible = False
   'Me.img_90.Visible = False
   'Me.img_95.Visible = False
   'Me.img_100.Visible = False
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub lv_lista_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      var_global_unidad = Me.lv_lista.selectedItem
      var_global_nombre_unidad = Me.lv_lista.selectedItem.SubItems(1)
      var_global_volumen_unidad = Me.lv_lista.selectedItem.SubItems(2)
      Me.txt_clave_unidad = var_global_unidad
      Me.txt_nombre_unidad = var_global_nombre_unidad
      Me.txt_volumen_unidad = var_global_volumen_unidad
      If CDbl(Me.txt_volumen_unidad) > 0 Then
         var_porcentaje = (CDbl(Me.txt_volumen_carga) * 100) / CDbl(Me.txt_volumen_unidad)
         Me.txt_porcentaje_carga = Round(var_porcentaje, 2)
      Else
         var_porcentaje = 0
         Me.txt_porcentaje_carga = Round(var_porcentaje, 2)
      End If
      
      
      rs.Open "update TB_ORACLE_GRUPOS_EMBARQUES set unidad = '" + var_global_unidad + "' where grupo = '" + Me.txt_embarque + "'", cnn, adOpenDynamic, adLockOptimistic
      Me.frm_lista.Visible = False
   End If
   If KeyAscii = 27 Then
      Me.frm_lista.Visible = False
   End If
End Sub

Private Sub lv_lista_LostFocus()
   Me.frm_lista.Visible = False
End Sub

Private Sub txt_clave_unidad_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      rsaux15.Open "select * from tb_usuarios where vcha_usu_usuario_id = '" + var_clave_usuario_global + "' and INTE_USU_PERMISO_CAMBIAR_TRANSPORTE = 1", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux15.EOF Then
         Me.lv_lista.ListItems.Clear
         rs.Open "select * from tb_oracle_Transportes order by nombre", cnn, adOpenDynamic, adLockOptimistic
         While Not rs.EOF
               Set list_item = lv_lista.ListItems.Add(, , rs!CLAVE)
               list_item.SubItems(1) = rs!nombre
               list_item.SubItems(2) = IIf(IsNull(rs!VOLUMEN), 0, rs!VOLUMEN)
               rs.MoveNext
         Wend
         rs.Close
         Me.frm_lista.Visible = True
      Else
         MsgBox "No tiene permitido seleccionar la unidad", vbOKOnly, "ATENCION"
         
      End If
      rsaux15.Close
   End If
End Sub

Private Sub txt_embarque_buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      rs.Open "select * FROM  TB_ORACLE_GRUPOS_EMBARQUES where embarque = '" + Me.txt_embarque_buscar + "'", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         MsgBox "El embarque se encuentra en el grupo " + rs!grupo, vbOKOnly, "ATENCION"
      Else
         MsgBox "El embarque no existe o no se a agrupado", vbOKOnly, "ATENCION"
      End If
      rs.Close
      Me.txt_embarque.SetFocus
   End If
   If KeyAscii = 27 Then
      Me.txt_embarque.SetFocus
   End If
End Sub

Private Sub txt_embarque_buscar_LostFocus()
   Me.frm_embarque.Visible = False
End Sub

Private Sub txt_embarque_Change()
   
   Me.txt_clave_unidad = ""
   Me.txt_nombre_unidad = ""
   Me.txt_volumen_unidad = ""
   Me.txt_volumen_carga = ""
   Me.txt_porcentaje_carga = ""
   Me.txt_volumen_carga_lectores = ""
   Me.txt_porcentaje_carga_lectores = ""
   Me.txt_embarques = ""
   Me.ProgressBar1.Value = 0
   Me.ProgressBar2.Value = 0
   'Me.img_0.Visible = True
   'Me.img_5.Visible = False
   'Me.img_10.Visible = False
   'Me.img_15.Visible = False
   'Me.img_20.Visible = False
   'Me.img_25.Visible = False
   'Me.img_30.Visible = False
   'Me.img_35.Visible = False
   'Me.img_40.Visible = False
   'Me.img_45.Visible = False
   'Me.img_50.Visible = False
   'Me.img_55.Visible = False
   'Me.img_60.Visible = False
   'Me.img_65.Visible = False
   'Me.img_70.Visible = False
   'Me.img_75.Visible = False
   'Me.img_80.Visible = False
   'Me.img_85.Visible = False
   'Me.img_90.Visible = False
   'Me.img_95.Visible = False
   'Me.img_100.Visible = False
   
End Sub

Private Sub txt_embarque_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      frmoracle_union_embarques.Show 1
      Call Command4_Click
   End If
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Command4_Click
   End If
End Sub

Private Sub txt_embarques_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 Then
      Me.frm_embarque.Visible = True
      Me.txt_embarque_buscar.SetFocus
   End If
End Sub
