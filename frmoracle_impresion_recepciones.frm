VERSION 5.00
Begin VB.Form frmoracle_impresion_recepciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Código a imprimir"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_codigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3945
   End
End
Attribute VB_Name = "frmoracle_impresion_recepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report
Dim comandoORA As New adodb.Command
Dim parametro As adodb.Parameter

Private Sub Form_Load()
   Top = 3000
   Left = 3550
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
   If Me.txt_codigo <> "" Then
      If KeyAscii = 13 Then
         Me.txt_codigo = UCase(Me.txt_codigo)
         rs.Open "select * From xxvia_vw_recepion_sub_ubi where folio_recepcion = '" + Me.txt_codigo + "'", cnnoracle_4, adOpenDynamic, adLockOptimistic
         If Not rs.EOF Then
            cnn.BeginTrans
            rsaux.Open "SELECT MAX(INTE_TEM_CONSECUTIVO) FROM TB_TEMP_ORACLE_REPORTE_RECEPCIONES", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux.EOF Then
               var_consecutivo = IIf(IsNull(rsaux(0).Value), 0, rsaux(0).Value)
            Else
               var_consecutivo = 0
            End If
            rsaux.Close
            var_consecutivo = var_consecutivo + 1
            rsaux.Open "insert into TB_TEMP_ORACLE_REPORTE_RECEPCIONES (inte_Tem_Consecutivo) values (" + CStr(var_consecutivo) + ")", cnn, adOpenDynamic, adLockOptimistic
            cnn.CommitTrans
            While Not rs.EOF
                  rsaux.Open "INSERT INTO TB_TEMP_ORACLE_REPORTE_RECEPCIONES (INTE_TEM_CONSECUTIVO, ALMACEN, CODIGO, DESCRIPCION, UBICACION, CANTIDAD, FOLIO) VALUES (" + CStr(var_consecutivo) + ",'" + rs!SUBINVENTORY_CODE + "','" + rs!CODIGO + "','" + rs!NOM_ARTICULO + "','" + IIf(IsNull(rs!UBI1), "", rs!UBI1) + "'," + CStr(rs!Cantidad) + ",'" + rs!FOLIO_RECEPCION + "')", cnn, adOpenDynamic, adLockOptimistic
                  rs.MoveNext
            Wend
                        
            Set reporte = appl.OpenReport(App.Path + "\rep_oracle_impresion_recepciones.rpt")
            reporte.RecordSelectionFormula = "{VW_ORACLE_IMPRESION_RECEPCIONES.INTE_TEM_CONSECUTIVO} = " + CStr(var_consecutivo)
            frmvistasprevias.cr.ReportSource = reporte
            For ntablas = 1 To reporte.Database.Tables.Count
                reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_reportes, parametros(4), parametros(5)
            Next ntablas
            frmvistasprevias.cr.ViewReport
            frmvistasprevias.Caption = "Recepción " + Me.txt_codigo
            frmvistasprevias.Show 1
            Set reporte = Nothing
                        
            
            rsaux.Open "delete from TB_TEMP_ORACLE_REPORTE_RECEPCIONES where inte_Tem_consecutivo = " + CStr(var_consecutivo), cnn, adOpenDynamic, adLockOptimistic
         Else
            MsgBox "La recepción no existe", vbOKOnly, "ATENCION"
         End If
         rs.Close
      End If
   End If
End Sub
