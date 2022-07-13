VERSION 5.00
Begin VB.Form frmreporte_mercadotecnia_ventas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte ventas"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "VC"
      Height          =   315
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Reporte de establecimientos de Vianney Catalog."
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame frm_mensaje 
      Height          =   735
      Left            =   60
      TabIndex        =   8
      Top             =   435
      Width           =   2805
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Espere un momento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   420
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Procesando Información "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   60
         TabIndex        =   9
         Top             =   150
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -150
      TabIndex        =   5
      Top             =   345
      Width           =   3135
   End
   Begin VB.CommandButton cmd_salir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2565
      Picture         =   "frmreporte_mercadotecnia_ventas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmd_imprimir 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Picture         =   "frmreporte_mercadotecnia_ventas.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir Alt + I"
      Top             =   30
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   " Periodo "
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   435
      Width           =   2805
      Begin VB.TextBox txt_mes 
         Height          =   315
         Left            =   1965
         TabIndex        =   2
         Top             =   270
         Width           =   705
      End
      Begin VB.TextBox txt_año 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Left            =   1545
         TabIndex        =   7
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   330
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmreporte_mercadotecnia_ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report

Private Sub cmd_imprimir_Click()
    Dim var_consecutivo As Double
    Dim var_año As Integer
    cnn.CommandTimeout = 360
    If IsNumeric(txt_año) Then
       If IsNumeric(txt_mes) Then
          If CInt(txt_mes) >= 1 Or CInt(txt_mes) <= 12 Then
             var_si = MsgBox("¿Desea el reporte por cliente?", vbYesNo, "ATENCION")
             var_si = 7
             If var_si = 6 Then
                frm_mensaje.Visible = True
                Me.Refresh
                Me.frm_mensaje.Refresh
                cnn.BeginTrans
                rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_TIENDA", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   var_consecutivo = 1
                Else
                   var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                End If
                If rsaux.State = 1 Then
                   rsaux.Close
                End If
                rsaux.Open "insert into TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_TIENDA (inte_tem_consecutivo, inte_tem_año, inte_tem_mes) values (" + CStr(var_consecutivo) + ", " + txt_año + ", " + txt_mes + ")", cnn, adOpenDynamic, adLockOptimistic
                rs.Close
                cnn.CommitTrans
                rs.Open "EXEC SP_REPORTE_MERCADOTECNIA_VENTAS_CLIENTE " + CStr(var_consecutivo) + "," + txt_año + "," + txt_mes, cnn, adOpenDynamic, adLockOptimistic
                rsaux8.Open "select  DISTINCT vcha_can_canal_venta, vcha_can_nombre from TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_TIENDA", cnn, adOpenDynamic, adLockOptimistic
                While Not rsaux8.EOF
                      Set reporte = appl.OpenReport(App.Path + "\rep_venta_mercadotecnia_tienda.rpt")
                      reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_TIENDA.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and not IsNull({VW_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_tienda.VCHA_CAN_CANAL_VENTA}) and {VW_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_TIENDA.vcha_can_canal_Venta} = '" + IIf(IsNull(rsaux8!VCHA_CAN_CANAL_VENTA), "", rsaux8!VCHA_CAN_CANAL_VENTA) + "'"
                      For ntablas = 1 To reporte.Database.Tables.Count
                          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                      Next ntablas
                      reporte.ExportOptions.FormatType = crEFTExcel80
                      reporte.ExportOptions.DestinationType = crEDTDiskFile
                      archivo = "c:\reportessid\ventas_mercadotecnia_cliente_" + Trim(rsaux8!vcha_can_nombre) + "_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                      reporte.ExportOptions.DiskFileName = archivo
                      reporte.Export False
                      Set reporte = Nothing
                      rsaux8.MoveNext
                Wend
                rsaux8.Close
                rs.Open "delete from TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_TIENDA where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                MsgBox "Se a terminado de guardar el archivo " + archivo
                Me.frm_mensaje.Visible = False
             Else
                frm_mensaje.Visible = True
                Me.Refresh
                Me.frm_mensaje.Refresh
                cnn.BeginTrans
                rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL", cnn, adOpenDynamic, adLockOptimistic
                If rs.EOF Then
                   var_consecutivo = 1
                Else
                   var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
                End If
                If rsaux.State = 1 Then
                   rsaux.Close
                End If
                rsaux.Open "insert into TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL (inte_tem_consecutivo, inte_tem_año, inte_tem_mes) values (" + CStr(var_consecutivo) + ", " + txt_año + ", " + txt_mes + ")", cnn, adOpenDynamic, adLockOptimistic
                rs.Close
                cnn.CommitTrans
                cnn.CommandTimeout = 360
                rs.Open "EXEC SP_REPORTE_MERCADOTECNIA_VENTAS " + CStr(var_consecutivo) + "," + txt_año + "," + txt_mes, cnn, adOpenDynamic, adLockOptimistic
             
                Set reporte = appl.OpenReport(App.Path + "\rep_venta_mercadotecnia.rpt")
                reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " and not IsNull({VW_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL.VCHA_CAN_CANAL_VENTA})"
                For ntablas = 1 To reporte.Database.Tables.Count
                   reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
                Next ntablas
                reporte.ExportOptions.FormatType = crEFTExcel80
                reporte.ExportOptions.DestinationType = crEDTDiskFile
                archivo = "c:\reportessid\ventas_mercadotecnia_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                reporte.ExportOptions.DiskFileName = archivo
                reporte.Export False
                Set reporte = Nothing
                rs.Open "delete from TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
                 MsgBox "Se a terminado de guardar el archivo " + archivo
                Me.frm_mensaje.Visible = False
             End If
          Else
             MsgBox "Mes incorrecto", vbOKOnly, "ATENCION"
          End If
       Else
          MsgBox "Mes incorrecto", vbOKOnly, "ATENCION"
       End If
    Else
       MsgBox "Año incorrecto", vbOKOnly, "ATENCION"
    End If
End Sub

Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
    Dim var_consecutivo As Double
    Dim var_año As Integer
    cnn.CommandTimeout = 360
    If IsNumeric(txt_año) Then
       If IsNumeric(txt_mes) Then
          If CInt(txt_mes) >= 1 Or CInt(txt_mes) <= 12 Then
             frm_mensaje.Visible = True
             Me.Refresh
             Me.frm_mensaje.Refresh
             cnn.BeginTrans
             rs.Open "select max(inte_tem_consecutivo) from TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_ESTABLECIMIENTO", cnn, adOpenDynamic, adLockOptimistic
             If rs.EOF Then
                var_consecutivo = 1
             Else
                var_consecutivo = IIf(IsNull(rs(0).Value), 0, rs(0).Value) + 1
             End If
             If rsaux.State = 1 Then
                rsaux.Close
             End If
             rsaux.Open "insert into TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_ESTABLECIMIENTO (inte_tem_consecutivo, inte_tem_año, inte_tem_mes) values (" + CStr(var_consecutivo) + ", " + txt_año + ", " + txt_mes + ")", cnn, adOpenDynamic, adLockOptimistic
             rs.Close
             cnn.CommitTrans
             rs.Open "EXEC SP_REPORTE_MERCADOTECNIA_VENTAS_ESTABLECIMIENTO " + CStr(var_consecutivo) + "," + txt_año + "," + txt_mes, cnn, adOpenDynamic, adLockOptimistic
             
             Set reporte = appl.OpenReport(App.Path + "\rep_venta_mercadotecnia_establecimiento.rpt")
             reporte.RecordSelectionFormula = "{VW_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_ESTABLECIMIENTO.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo) + " AND {VW_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_ESTABLECIMIENTO.VCHA_cAN_cANAL_vENTA} = '10'"
             For ntablas = 1 To reporte.Database.Tables.Count
                 reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
             Next ntablas
             reporte.ExportOptions.FormatType = crEFTExcel80
             reporte.ExportOptions.DestinationType = crEDTDiskFile
             archivo = "c:\reportessid\ventas_mercadotecnia_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
             reporte.ExportOptions.DiskFileName = archivo
             reporte.Export False
             Set reporte = Nothing
              
             rs.Open "delete from TB_TEMP_REPORTE_MERCADOTECNIA_VENTAS_MENSUAL_ESTABLECIMIENTO where inte_tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
             MsgBox "Se a terminado de guardar el archivo " + archivo
             Me.frm_mensaje.Visible = False
          Else
             MsgBox "Mes incorrecto", vbOKOnly, "ATENCION"
          End If
       Else
          MsgBox "Mes incorrecto", vbOKOnly, "ATENCION"
       End If
    Else
       MsgBox "Año incorrecto", vbOKOnly, "ATENCION"
    End If

End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   frm_mensaje.Visible = False
   If var_empresa = "03" Then
      Command1.Visible = True
   Else
      Command1.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_reporte_valuacion_devoluciones)
End Sub

Private Sub txt_año_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.txt_mes.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub

Private Sub txt_mes_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
   Case 48 To 57, 52, 13, 8, 46
   Case Else
       KeyAscii = 0
   End Select
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
   If KeyAscii = 27 Then
      Unload Me
   End If
End Sub
