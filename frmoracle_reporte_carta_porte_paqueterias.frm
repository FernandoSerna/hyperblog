VERSION 5.00
Begin VB.Form frmoracle_reporte_carta_porte_paqueterias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Embarque - Paqueterias"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3195
      Begin VB.TextBox txt_embarque 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1125
         TabIndex        =   1
         Top             =   180
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Embarque:"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   345
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmoracle_reporte_carta_porte_paqueterias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
   Call activa_forma(var_activa_forma_existencias_generales)
End Sub

Private Sub txt_embarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If IsNumeric(Me.txt_embarque) Then
          rs.Open "select * from xxvia_tb_encabezado_embarques where embarque = " + Me.txt_embarque, cnnoracle_4, adOpenDynamic, adLockOptimistic
          If Not rs.EOF Then
             VAR_ESTATUS = IIf(IsNull(rs!char_emb_estatus), "", rs!char_emb_estatus)
             If VAR_ESTATUS = "I" Then
                var_cadena = "select inte_emb_embarque EMBARQUE, source_header_number PEDIDO, a.segment1 CODIGO_INTERNO, c.CLASIFICACIONSAT CODIGO_SAT, c.DESCRIPTION, c.UOM_SAT ,c.UNIT_WEIGHT PESO, sum(floa_sal_cantidad_leida) as CANTIDAD, sum(floa_sal_cantidad_leida  * c.UNIT_WEIGHT) PESO_TOTAL from xxvia_tb_salidas_cajas a, xxvia_tb_encabezado_embarques b, xxvia_system_items_b c where a.inte_emb_embarque = b.embarque and organizacion = organization_id and a.segment1 = c.segment1  and embarque = " + Me.txt_embarque + " and floa_sal_cantidad_leida>0 group by inte_emb_embarque, a.segment1, c.DESCRIPTION, c.CLASIFICACIONSAT, c.UOM_SAT ,c.UNIT_WEIGHT, source_header_number"
                Set oexcel = CreateObject("Excel.Application")
                Set owbook = oexcel.Workbooks.Add
                Set osheet = owbook.Worksheets(1)
                osheet.Name = Me.txt_embarque
                Screen.MousePointer = vbHourglass
                iFila = 1
                ifila2 = 1
                icol2 = 1
                iCol = 1
                var_cadena = "select inte_emb_embarque EMBARQUE, SOURCE_HEADER_NUMBER PEDIDO,a.segment1 CODIGO_INTERNO, c.CLASIFICACIONSAT CODIGO_SAT, c.DESCRIPTION, c.UOM_SAT ,c.UNIT_WEIGHT PESO, sum(floa_sal_cantidad_leida) as CANTIDAD, sum(floa_sal_cantidad_leida  * c.UNIT_WEIGHT) PESO_TOTAL from xxvia_tb_salidas_cajas a, xxvia_tb_encabezado_embarques b, xxvia_system_items_b c where a.inte_emb_embarque = b.embarque and organizacion = organization_id and a.segment1 = c.segment1  and embarque = " + Me.txt_embarque + " and floa_sal_cantidad_leida>0 group by inte_emb_embarque, a.segment1, c.DESCRIPTION, c.CLASIFICACIONSAT, c.UOM_SAT ,c.UNIT_WEIGHT, SOURCE_HEADER_NUMBER ORDER BY SOURCE_HEADER_NUMBER"
                If rsaux10.State = 1 Then
                   rsaux10.Close
                End If
                rsaux10.Open var_cadena, cnnoracle_4, adOpenDynamic, adLockOptimistic
                For i = 0 To rsaux10.Fields.Count - 1
                    osheet.Cells(iFila, i + 1) = rsaux10.Fields(i).Name
                Next
                iFila = iFila + 1
                With osheet
                     .Cells(iFila, iCol).CopyFromRecordset rsaux10
                End With
                archivo = "c:\reportessid\EMBARQUE_" + Me.txt_embarque + "_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
                owbook.SaveAs archivo
                oexcel.Visible = True
                Set oexcel = Nothing
                Screen.MousePointer = vbDefault
                rsaux10.Close
             Else
                MsgBox "El embarque no ha sido cerrado aun.", vbOKOnly, "ATENCION"
             End If
          Else
             MsgBox "El embarque no existe", vbOKOnly, "ATENCION"
          End If
          rs.Close
       Else
          MsgBox "Número de embarque incorrecto", vbOKOnly, "ATENCION"
       End If
    End If
End Sub
