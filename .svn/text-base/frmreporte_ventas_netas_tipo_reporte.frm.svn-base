VERSION 5.00
Begin VB.Form frmreporte_ventas_netas_tipo_reporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de Reporte de Ventas Netas"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_salir 
      Height          =   330
      Left            =   3210
      Picture         =   "frmreporte_ventas_netas_tipo_reporte.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   15
      Width           =   345
   End
   Begin VB.CommandButton cmd_aceptar 
      Height          =   330
      Left            =   60
      Picture         =   "frmreporte_ventas_netas_tipo_reporte.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   345
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   15
      TabIndex        =   6
      Top             =   345
      Width           =   3630
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tipo de Reporte "
      Height          =   2070
      Left            =   75
      TabIndex        =   0
      Top             =   450
      Width           =   3510
      Begin VB.OptionButton opt_grupo 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1326
         Width           =   1545
      End
      Begin VB.OptionButton opt_titular 
         Caption         =   "Titular"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1650
         Width           =   1545
      End
      Begin VB.OptionButton opt_ruta 
         Caption         =   "Ruta"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1004
         Width           =   1545
      End
      Begin VB.OptionButton opt_agente 
         Caption         =   "Agente"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   682
         Width           =   1545
      End
      Begin VB.OptionButton opt_canal_venta 
         Caption         =   "Canal de Venta"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmreporte_ventas_netas_tipo_reporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

End Sub

Private Sub cmd_aceptar_Click()
   If Me.opt_canal_venta.Value = True Then
      var_activa_forma_salidas = "frmreporte_ventas_netas_tipo_reporte"
      frmreporte_ventas_netas_canal_venta.Show
      Me.Enabled = False
   End If
   If Me.opt_agente.Value = True Then
      var_activa_forma_reporte_valuacion_devoluciones = "frmreporte_ventas_netas_tipo_reporte"
      frmreporte_ventas_netas.Show
      Me.Enabled = False
   End If
   If Me.opt_ruta.Value = True Then
      var_activa_forma_salidas = "frmreporte_ventas_netas_tipo_reporte"
      frmreporte_ventas_netas_ruta.Show
      Me.Enabled = False
   End If
   If Me.opt_titular.Value = True Then
      var_activa_forma_salidas = "frmreporte_ventas_netas_tipo_reporte"
      frmreporte_ventas_netas_titular.Show
      Me.Enabled = False
   End If
   If Me.opt_grupo.Value = True Then
      var_activa_forma_salidas = "frmreporte_ventas_netas_tipo_reporte"
      frmreporte_ventas_netas_grupo.Show
      Me.Enabled = False
   End If
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
Dim dl As Long                                 ' Valor devuelto por la función API
Dim sAttributes As String                  ' Aributos
Dim sDriver As String                       ' Nombre del controlador
Dim sDescription As String                ' Descripción del DSN
Dim sDsnName As String                  ' Nombre del DSN

   cnn.Close
   cnn.Open var_conexion_string_distribucion

   Const ODBC_ADD_SYS_DSN As Long = 4         ' Se creará un DSN de sistema
   Const vbAPINull As Long = 0&                         ' Puntero NULL

   ' se elimina
   Const ODBC_REMOVE_SYS_DSN As Long = 6    ' Se eliminará un DSN de sistema
   sDsnName = "DSN=sqlsistema"
   sDriver = "SQL Server"
   dl = SQLConfigDataSource(0&, ODBC_REMOVE_SYS_DSN, sDriver, sDsnName)

   'se crea
   sDsnName = "sqlsistema"
   sDescription = "sqlsistema"
   sDriver = "SQL Server"
   sAttributes = "DSN=" & sDsnName & Chr(0)
   sAttributes = sAttributes & "Server=" + parametros(0) & Chr$(0)
   sAttributes = sAttributes & "Description=" & sDescription & Chr(0)
   sAttributes = sAttributes & "Database=" + var_bd_reportes & Chr(0)
   strAttributes = strAttributes & "UID=sa" & Chr$(0)
   strAttributes = strAttributes & "PWD=elia" & Chr$(0)
   dl = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, sDriver, sAttributes)
   
   
   Top = 2500
   Left = 4000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_packing_list)
End Sub

