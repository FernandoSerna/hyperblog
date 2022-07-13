VERSION 5.00
Begin VB.Form frmejecuta_calculo_pagos_multibondeados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecuta calculo de pagos de multibondeados"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Periodo "
      Height          =   720
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4815
      Begin VB.CommandButton cmd_aceptar_pedidos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4380
         Picture         =   "frmejecuta_calculo_pagos_multibondeados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aceptar Alt + A"
         Top             =   225
         Width           =   330
      End
      Begin VB.ComboBox com_mes 
         Height          =   315
         ItemData        =   "frmejecuta_calculo_pagos_multibondeados.frx":014A
         Left            =   480
         List            =   "frmejecuta_calculo_pagos_multibondeados.frx":0172
         TabIndex        =   2
         Top             =   225
         Width           =   2190
      End
      Begin VB.TextBox txt_año 
         Height          =   300
         Left            =   3405
         TabIndex        =   1
         Top             =   232
         Width           =   945
      End
      Begin VB.Label Mes 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   75
         TabIndex        =   4
         Top             =   285
         Width           =   300
      End
      Begin VB.Label Label1 
         Caption         =   "Año"
         Height          =   195
         Left            =   2790
         TabIndex        =   3
         Top             =   285
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmejecuta_calculo_pagos_multibondeados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_pedidos_Click()
   If Me.txt_año = 2005 Or Me.txt_año = 2006 Or Me.txt_año = 2007 Or Me.txt_año = 2008 Or Me.txt_año = 2009 Or Me.txt_año = 2010 Or Me.txt_año = 2011 Or Me.txt_año = 2012 Or Me.txt_año = 2013 Or Me.txt_año = 2014 Then
      If com_mes = "ENERO" Then
         var_mes = 1
         var_año = CInt(txt_año)
         var_mes_str = "Enero"
      End If
      If com_mes = "FEBRERO" Then
         var_mes = 2
         var_año = CInt(txt_año)
         var_mes_str = "Febrero"
      End If
      If com_mes = "MARZO" Then
         var_mes = 3
         var_año = CInt(txt_año)
         var_mes_str = "Marzo"
      End If
      If com_mes = "ABRIL" Then
         var_mes = 4
         var_año = CInt(txt_año)
         var_mes_str = "Abril"
      End If
      If com_mes = "MAYO" Then
         var_mes = 5
         var_año = CInt(txt_año)
         var_mes_str = "Mayo"
      End If
      If com_mes = "JUNIO" Then
         var_mes = 6
         var_año = CInt(txt_año)
         var_mes_str = "Junio"
      End If
      If com_mes = "JULIO" Then
         var_mes = 7
         var_año = CInt(txt_año)
         var_mes_str = "Julio"
      End If
      If com_mes = "AGOSTO" Then
         var_mes = 8
         var_año = CInt(txt_año)
         var_mes_str = "Agosto"
      End If
      If com_mes = "SEPTIEMBRE" Then
         var_mes = 9
         var_año = CInt(txt_año)
         var_mes_str = "Septiembre"
      End If
      If com_mes = "OCTUBRE" Then
         var_mes = 10
         var_año = CInt(txt_año)
         var_mes_str = "Octubre"
      End If
      If com_mes = "NOVIEMBRE" Then
         var_mes = 11
         var_año = CInt(txt_año)
         var_mes_str = "Noviembre"
      End If
      If com_mes = "DICIEMBRE" Then
         var_mes = 12
         var_año = CInt(txt_año) + 1
         var_mes_str = "Diciembre"
      End If
      var_si = MsgBox("Se correran los calculos correspondientes al periodo de " + var_mes_str + " del " + txt_año, vbYesNo, "ATENCION")
      If var_si = 6 Then
         var_si = MsgBox("Confirmar la ejecución de los calculos del periodo de " + var_mes_str + " del " + txt_año, vbYesNo, "ATENCION")
         If var_si = 6 Then
            cnn.CommandTimeout = 360
            rs.Open "DELETE FROM TB_IMPORTE_PAGOS_MULTIBONDEADOS WHERE INTE_DMU_AÑO = " + Me.txt_año + " AND INTE_DMU_MES = " + CStr(var_mes), cnn, adOpenDynamic, adLockOptimistic
            rs.Open "exec SP_CALCULO_PAGOS_MULTIBONDEADOS " + Me.txt_año + "," + CStr(var_mes), cnn, adOpenDynamic, adLockOptimistic
            MsgBox "Se a terminado de ejecutar el calculo", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "Año incorrecto", vbOKOnly, "ATENCION"
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
   Call activa_forma(var_activa_forma_packing_list)
End Sub
