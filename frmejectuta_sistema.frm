VERSION 5.00
Begin VB.Form frmejectuta_sistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración para la actualización de la versión del sistema"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6780
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6360
      Picture         =   "frmejectuta_sistema.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmejectuta_sistema.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   15
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   60
      TabIndex        =   5
      Top             =   360
      Width           =   6705
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   6675
      Begin VB.TextBox txt_ruta_actualizacion 
         Height          =   315
         Left            =   1575
         TabIndex        =   4
         Top             =   600
         Width           =   4650
      End
      Begin VB.TextBox txt_ruta_local 
         Height          =   315
         Left            =   1575
         TabIndex        =   3
         Top             =   255
         Width           =   4650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Actualización:"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   660
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Local:"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmejectuta_sistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection
Dim var_fecha_local As Date
Dim var_fecha_servidor As Date
Dim var_archivo_local As String
Dim var_archivo_servidor As String
Dim var_ruta As String

Private Sub cmd_aceptar_Click()
   Dim var_nueva_ruta_local As String
   Dim var_nueva_ruta_actualizacion As String
   Set var_tabla = CreateObject("ADODB.connection")
   var_tabla.CursorLocation = adUseClient
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   var_ruta = App.Path
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   rs.Open "update sistema set ruta_local = '" + txt_ruta_local + "', ruta_servi = '" + txt_ruta_actualizacion + "'", var_tabla, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   var_cadena_seguridad = ""
   Top = 3000
   Left = 3000
   Set var_tabla = CreateObject("ADODB.connection")
   var_tabla.CursorLocation = adUseClient
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   var_ruta = App.Path
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + var_ruta + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   rs.Open "select * from sistema", var_tabla, adOpenDynamic, adLockOptimistic
   If IsNull(rs!ruta_local) = False Then
      If IsNull(rs!ruta_servi) = False Then
         var_archivo_local = Trim(rs!ruta_local)
         var_archivo_servidor = Trim(rs!ruta_servi)
         txt_ruta_local = var_archivo_local
         txt_ruta_actualizacion = var_archivo_servidor
      End If
   End If
   rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   var_swpassword = False
   var_modifica_registro = False
   Call activa_forma(var_activa_forma_ejecuta_sistema)
End Sub
