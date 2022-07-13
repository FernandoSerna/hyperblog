VERSION 5.00
Begin VB.Form frmejecuta_calculos_Mensuales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculos Mensuales"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_crea_tabla_multibondeados 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   435
      Picture         =   "frmejecuta_calculos_Mensuales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Crear tabla de descuentos para multibondeados"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4320
      Picture         =   "frmejecuta_calculos_Mensuales.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   30
      Width           =   330
   End
   Begin VB.CommandButton cmd_aceptar_pedidos 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Picture         =   "frmejecuta_calculos_Mensuales.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aceptar Alt + A"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame x 
      Caption         =   " Periodo "
      Height          =   930
      Left            =   120
      TabIndex        =   0
      Top             =   465
      Width           =   4530
      Begin VB.TextBox txt_año 
         Height          =   300
         Left            =   3465
         TabIndex        =   4
         Top             =   390
         Width           =   945
      End
      Begin VB.ComboBox com_mes 
         Height          =   315
         ItemData        =   "frmejecuta_calculos_Mensuales.frx":0886
         Left            =   540
         List            =   "frmejecuta_calculos_Mensuales.frx":08AE
         TabIndex        =   2
         Top             =   390
         Width           =   2190
      End
      Begin VB.Label Label1 
         Caption         =   "Año"
         Height          =   195
         Left            =   2850
         TabIndex        =   3
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Mes 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   405
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   90
      TabIndex        =   6
      Top             =   330
      Width           =   4575
   End
End
Attribute VB_Name = "frmejecuta_Calculos_Mensuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If Me.txt_año = 2005 Or Me.txt_año = 2006 Or Me.txt_año = 2007 Or Me.txt_año = 2008 Or Me.txt_año = 2009 Or Me.txt_año = 2010 Or Me.txt_año = 2011 Or Me.txt_año = 2012 Or Me.txt_año = 2013 Or Me.txt_año = 2014 Then
      If com_mes = "ENERO" Then
         var_mes = 2
         var_año = CInt(txt_año)
         var_mes_str = "Enero"
      End If
      If com_mes = "FEBRERO" Then
         var_mes = 3
         var_año = CInt(txt_año)
         var_mes_str = "Febrero"
      End If
      If com_mes = "MARZO" Then
         var_mes = 4
         var_año = CInt(txt_año)
         var_mes_str = "Marzo"
      End If
      If com_mes = "ABRIL" Then
         var_mes = 5
         var_año = CInt(txt_año)
         var_mes_str = "Abril"
      End If
      If com_mes = "MAYO" Then
         var_mes = 6
         var_año = CInt(txt_año)
         var_mes_str = "Mayo"
      End If
      If com_mes = "JUNIO" Then
         var_mes = 7
         var_año = CInt(txt_año)
         var_mes_str = "Junio"
      End If
      If com_mes = "JULIO" Then
         var_mes = 8
         var_año = CInt(txt_año)
         var_mes_str = "Julio"
      End If
      If com_mes = "AGOSTO" Then
         var_mes = 9
         var_año = CInt(txt_año)
         var_mes_str = "Agosto"
      End If
      If com_mes = "SEPTIEMBRE" Then
         var_mes = 10
         var_año = CInt(txt_año)
         var_mes_str = "Septiembre"
      End If
      If com_mes = "OCTUBRE" Then
         var_mes = 11
         var_año = CInt(txt_año)
         var_mes_str = "Octubre"
      End If
      If com_mes = "NOVIEMBRE" Then
         var_mes = 12
         var_año = CInt(txt_año)
         var_mes_str = "Noviembre"
      End If
      If com_mes = "DICIEMBRE" Then
         var_mes = 1
         var_año = CInt(txt_año) + 1
         var_mes_str = "Diciembre"
      End If
      var_si = MsgBox("Se correran los calculos correspondientes al periodo de " + var_mes_str + " del " + txt_año, vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la ejecución de los calculos del periodo de " + var_mes_str + " del " + txt_año, vbYesNo, "ATENCION")
         If var_si = 6 Then
            If var_mes < 10 Then
               VAR_MES_FECHA = "0" + Trim(CStr(var_mes))
            Else
               VAR_MES_FECHA = Trim(CStr(var_mes))
            End If
            var_fecha = "{d '" + CStr(var_año) + "-" + VAR_MES_FECHA + "-01'}"
            MsgBox "EXEC SP_EJECUTA_CALCULOS " + var_fecha
            var_cadena = "EXEC SP_EJECUTA_CALCULOS " + var_fecha
            cnn.CommandTimeout = 360
            rs.Open var_cadena, cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a terminado de ejecutar los calculos", vbOKOnly, "ATENCION"
         End If
      End If
   End If
   
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub com_mes_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = 0
   End If
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   txt_año = Year(Date)
   var_mes = Month(Date)
   If var_mes = 1 Then
      com_mes = "ENERO"
   End If
   If var_mes = 2 Then
      com_mes = "FEBRERO"
   End If
   If var_mes = 3 Then
      com_mes = "MARZO"
   End If
   If var_mes = 4 Then
      com_mes = "ABRIL"
   End If
   If var_mes = 5 Then
      com_mes = "MAYO"
   End If
   If var_mes = 6 Then
      com_mes = "JUNIO"
   End If
   If var_mes = 7 Then
      com_mes = "JULIO"
   End If
   If var_mes = 8 Then
      com_mes = "AGOSTO"
   End If
   If var_mes = 9 Then
      com_mes = "SEPTIEMBRE"
   End If
   If var_mes = 10 Then
      com_mes = "OCTUBRE"
   End If
   If var_mes = 11 Then
      com_mes = "NOVIEMBRE"
   End If
   If var_mes = 12 Then
      com_mes = "DICIEMBRE"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_articulos2)
End Sub
