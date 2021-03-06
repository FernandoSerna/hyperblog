VERSION 5.00
Begin VB.Form frmConImpEtiPre 
   Caption         =   "Configuracion Impresora Etiquetas "
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4170
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3720
      Picture         =   "frmConImpEtiPre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   330
   End
   Begin VB.CommandButton cmd_guardar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "frmConImpEtiPre.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Guardar Alt + G"
      Top             =   0
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   120
      TabIndex        =   4
      Top             =   270
      Width           =   4050
   End
   Begin VB.Frame fraConfiguacionImpresion 
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.TextBox txt_nombreMaquina 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cbo_Puerto 
         Height          =   315
         ItemData        =   "frmConImpEtiPre.frx":073C
         Left            =   1560
         List            =   "frmConImpEtiPre.frx":0746
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbl_nomMaquina 
         Caption         =   "Nombre Impresora:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Puerto:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmConImpEtiPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnnCompucaja As New ADODB.Connection
Dim strNomMaquina As String

Private Sub cbo_Puerto_Click()
    If cbo_Puerto.Text = "- USB" Then
        txt_nombreMaquina.Visible = True
    Else
        txt_nombreMaquina.Text = ""
        txt_nombreMaquina.Visible = False
    End If

End Sub

Private Sub cmd_guardar_Click()
    Dim rsGuarda As New ADODB.recordSet

    If Conectar_BD(cnnCompucaja, "compucaja", "srvtdacantia") Then
        strNomMaquina = fun_NombrePc
        rsGuarda.Open "Select * " & _
                    "from tb_etiquetasConfiguracionImpresora  " & _
                    "where vcha_eci_maquina= '" & strNomMaquina & "'", _
                cnnCompucaja, _
                adOpenDynamic, _
                adLockOptimistic
        If rsGuarda.RecordCount > 0 Then
            rsGuarda.Close
            rsGuarda.Open "Update  tb_etiquetasConfiguracionImpresora " & _
                            "Set vcha_eci_puerto = '" & cbo_Puerto.Text & "', " & _
                                "vcha_eci_maquinaImpresora = '" & txt_nombreMaquina.Text & "' " & _
                            "where vcha_eci_maquina= '" & strNomMaquina & "'", _
                    cnnCompucaja, _
                    adOpenDynamic, _
                    adLockOptimistic
        Else
            rsGuarda.Close
            rsGuarda.Open "Insert Into  tb_etiquetasConfiguracionImpresora " & _
                            "(vcha_eci_puerto, " & _
                            "vcha_eci_maquinaImpresora, " & _
                            "vcha_eci_maquina) " & _
                        "Values " & _
                            "('" & cbo_Puerto.Text & "', " & _
                            " '" & txt_nombreMaquina.Text & "' )" & _
                            " '" & strNomMaquina & "'", _
                    cnnCompucaja, _
                    adOpenDynamic, _
                    adLockOptimistic
        
        End If
        cnnCompucaja.Close
        Set rsGuarda = Nothing
    Else
        MsgBox "Error al conectarse a compucaja", vbCritical, "SID"
    End If
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsEtiquetaConfiguracion As New ADODB.recordSet
    
    
    If Conectar_BD(cnnCompucaja, "compucaja", "srvtdacantia") Then
        strNomMaquina = fun_NombrePc
        lbl_nomMaquina.Caption = "Nom. Impresora"
        rsEtiquetaConfiguracion.Open "Select * " & _
                                    "from tb_etiquetasConfiguracionImpresora  " & _
                                    "where vcha_eci_maquina= '" & strNomMaquina & "'", _
                                cnnCompucaja, _
                                adOpenDynamic, _
                                adLockOptimistic
        If rsEtiquetaConfiguracion.RecordCount > 0 Then
            
            cbo_Puerto.Text = rsEtiquetaConfiguracion("vcha_eci_puerto").Value
            If cbo_Puerto.Text = "- USB" Then
                txt_nombreMaquina.Text = rsEtiquetaConfiguracion("vcha_eci_maquinaImpresora").Value
                txt_nombreMaquina.Visible = True
            Else
                txt_nombreMaquina.Text = ""
                txt_nombreMaquina.Visible = False
            End If
        End If
        rsEtiquetaConfiguracion.Close
    Else
        MsgBox "Error al conectarse a compucaja", vbCritical, "SID"
    End If
    Set rsEtiquetaConfiguracion = Nothing
    cnnCompucaja.Close
End Sub



Private Function Conectar_BD(ByRef cnnCBD As ADODB.Connection, ByVal bd As String, ByVal servidor As String) As Boolean
    'Variables de bloque
    Dim strConnection_String As String
    
On Error GoTo Error_Conectar_BDS
    Conectar_BD = True
    'Establecer connection strings para realizar las conexiones a las bases de
    'datos
    
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

