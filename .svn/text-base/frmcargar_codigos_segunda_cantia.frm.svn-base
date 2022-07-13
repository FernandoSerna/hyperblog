VERSION 5.00
Begin VB.Form frmcargar_codigos_segunda_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar códigos de segunda de cantia"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbusqueda_pedido 
      Height          =   3900
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   105
         TabIndex        =   5
         Top             =   510
         Width           =   3135
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   3330
         Pattern         =   "*.xls"
         TabIndex        =   4
         Top             =   510
         Width           =   3075
      End
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   90
         TabIndex        =   3
         Top             =   930
         Width           =   3150
      End
      Begin VB.CommandButton cmd_buscar_pedido 
         Caption         =   "Cargar archivo"
         Height          =   465
         Left            =   3330
         TabIndex        =   2
         Top             =   2805
         Width           =   3060
      End
      Begin VB.TextBox txt_ruta 
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3390
         Width           =   6315
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Caption         =   " Busqueda de archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   30
         TabIndex        =   6
         Top             =   120
         Width           =   6465
      End
   End
End
Attribute VB_Name = "frmcargar_codigos_segunda_cantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim appl As New CRAXDRT.Application
Dim reporte As New CRAXDRT.Report



Private Sub cmd_buscar_pedido_Click()
On Error GoTo SALIR:
   If Me.txt_ruta <> "" Then
      strConnectionString = "DRIVER=Microsoft Excel Driver (*.xls); DBQ=" & Me.txt_ruta
      rsaux2.Open "SELECT * FROM [hoja1$]", strConnectionString
      rsaux3.Open "SELECT MAX(INTE_PED_NUMERO) FROM TB_PEDIDOS_CANTIA", cnn, adOpenDynamic, adLockOptimistic
      If Not rsaux3.EOF Then
         var_numero_pedido = IIf(IsNull(rsaux3(0).Value), 0, rsaux3(0).Value) + 1
      Else
         var_numero_pedido = 1
      End If
      rsaux3.Close
      While Not rsaux2.EOF
            If Not IsNull(rsaux2!codigo) Then
               If rsaux2!pedido > 0 Then
                  rsaux3.Open "insert into TB_PEDIDOS_CANTIA (INTE_PED_NUMERO, VCHA_ART_ARTICULO_ID, FLOA_PED_CANTIDAD_PEDIDA, DTIM_PED_FECHA) values (" + CStr(var_numero_pedido) + ",'" + rsaux2!codigo + "'," + CStr(rsaux2!pedido) + ",GETDATE())", cnn, adOpenDynamic, adLockOptimistic
               End If
            End If
            rsaux2.MoveNext
      Wend
      rsaux2.Close
      
      rsaux2.Open "SELECT * FROM TB_PEDIDOS_CANTIA WHERE INTE_PED_NUMERO = " + CStr(var_numero_pedido), cnn, adOpenDynamic, adLockOptimistic
      While Not rsaux2.EOF
            If rsaux3.State = 1 Then
               rsaux3.Close
            End If
            rsaux3.Open "SELECT * FROM TB_EXISTENCIAS WHERE VCHA_ALM_ALMACEN_ID = 'PTVH' AND VCHA_aRT_ARTICULO_ID = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux3.EOF Then
               rsaux4.Open "UPDATE TB_PEDIDOS_CANTIA SET FLOA_EXI_CANTIDAD = " + CStr(IIf(IsNull(rsaux3!floa_Exi_Cantidad), 0, rsaux3!floa_Exi_Cantidad)) + " WHERE INTE_PED_NUMERO = " + CStr(var_numero_pedido) + " AND VCHA_ART_ARTICULO_ID = '" + rsaux3!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux3.Close
            rsaux3.Open "SELECT * FROM TB_CODIGOS_PROVEEDOR_CANTIA WHERE VCHA_ART_ARTICULO_ID = '" + rsaux2!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            If Not rsaux3.EOF Then
               rsaux4.Open "UPDATE TB_PEDIDOS_CANTIA SET VCHA_COD_CODIGO_BARRAS = '" + IIf(IsNull(rsaux3!VCHA_COD_CODIGO_BARRAS), "", rsaux3!VCHA_COD_CODIGO_BARRAS) + "', VCHA_COD_CODIGO_PROVEEDOR = '" + IIf(IsNull(rsaux3!VCHA_COD_CODIGO_PROVEEDOR), "", rsaux3!VCHA_COD_CODIGO_PROVEEDOR) + "', VCHA_COD_NOMBRE_PROVEEDOR = '" + IIf(IsNull(rsaux3!VCHA_COD_NOMBRE_PROVEEDOR), "", rsaux3!VCHA_COD_NOMBRE_PROVEEDOR) + "' WHERE INTE_PED_NUMERO = " + CStr(var_numero_pedido) + " AND VCHA_ART_ARTICULO_ID = '" + rsaux3!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
            End If
            rsaux3.Close
            If Mid(rsaux2!vcha_Art_articulo_id, 1, 1) = "T" Then
               rsaux3.Open "select * from tb_equivalencias where vcha_Art_articulo_id = '" + rsaux2!vcha_Art_articulo_id + "' and substring(vcha_equ_codigo_equivalente,1,5) = '64624'", cnn, adOpenDynamic, adLockOptimistic
               If Not rsaux3.EOF Then
                  rsaux4.Open "UPDATE TB_PEDIDOS_CANTIA SET VCHA_COD_CODIGO_BARRAS = '" + IIf(IsNull(rsaux3!vcha_equ_codigo_equivalente), "", rsaux3!vcha_equ_codigo_equivalente) + "' WHERE INTE_PED_NUMERO = " + CStr(var_numero_pedido) + " AND VCHA_ART_ARTICULO_ID = '" + rsaux3!vcha_Art_articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
               End If
               rsaux3.Close
            End If
            rsaux2.MoveNext
      Wend
      rsaux2.Close
      
      
      
      Set reporte = appl.OpenReport(App.Path + "\rep_pedido_Cantia.rpt")
      frmvistasprevias.cr.ReportSource = reporte
      reporte.RecordSelectionFormula = "{VW_PEDIDO_UBICACIONES.INTE_PED_NUMERO} = " + CStr(var_numero_pedido)
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_crystal, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de Movimientos"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("Desea exportar el archivo", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_pedido_Cantia_Excel.rpt")
         reporte.RecordSelectionFormula = "{VW_PEDIDO_UBICACIONES.INTE_PED_NUMERO} = " + CStr(var_numero_pedido)
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_crystal, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reporte_pedido_cantia_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         
'#######################################################################
'con este se abre el libre de excel
'########################################################################
Dim oExcel As Excel.Application
Dim oWorkBook As Excel.Workbook
Dim oSheet As Excel.Worksheet, var_ini, var_fin

Set oExcel = New Excel.Application

oExcel.Visible = True

    Set oWorkBook = oExcel.Workbooks.Open(archivo)

    Set oSheet = oExcel.Workbooks(1).Worksheets("rep_pedido_Cantia_Excel.rpt")

'####################################################################
'Con esto es como si se ejecutara una macro
'####################################################################
' DA FORMATO CONDICIONAL A TODO EL RANGO

With oExcel
    Range("A1").Select
    Selection.EntireRow.Delete
    Application.Goto Reference:="R2C5"
    Selection.NumberFormat = "m/d/yyyy"
    Application.Goto Reference:="R4C1"
    ActiveCell.Offset(0, 0).Range("A1:H471").Sort Key1:=ActiveCell.Offset(0, 0). _
        Range("A1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, _
        MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:= _
        xlSortTextAsNumbers
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Find(What:="t", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveCell.Range("A1:A5").Select
    Selection.EntireRow.Insert
    Application.Goto Reference:="R4C1"
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Font
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveCell.Select
    Ordena
    Application.Goto Reference:="R4C1"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Font
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveCell.Select
    Ordena
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    ActiveCell.Range("A1:A5").Select
    Selection.EntireRow.Insert
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlLeft).LineStyle = xlNone
    Selection.Borders(xlRight).LineStyle = xlNone
    Selection.Borders(xlTop).LineStyle = xlNone
    Selection.Borders(xlBottom).LineStyle = xlNone
    Application.Goto Reference:="R4C1"
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(-1, 0).Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.Goto Reference:="R4C1"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Columns.AutoFit
    ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.columnWidth = 78.86
    Range("A1").Select
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$4"
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.166645835937174)
        .RightMargin = Application.InchesToPoints(0.166645835937174)
        .TopMargin = Application.InchesToPoints(0.236775958560735)
        .BottomMargin = Application.InchesToPoints(0.236775958560735)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0.236775958560735)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = "Página &P de &N"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.166645835937174)
        .RightMargin = Application.InchesToPoints(0.166645835937174)
        .TopMargin = Application.InchesToPoints(0.236775958560735)
        .BottomMargin = Application.InchesToPoints(0.236775958560735)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0.236775958560735)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    ActiveWorkbook.Save
    
   
    
    '' fin de excel
End With


 'With oExcel
 '       .Cells.Select
 '       .Selection.Font.Name = "Arial"
 '       .Selection.Font.Size = 8
 '       .Selection.Interior.ColorIndex = 2
 '   End With



'#######################################################################
'con esto cierro el libro de excel
'#######################################################################

oExcel.DisplayAlerts = False
    oWorkBook.Save
    oWorkBook.Close SaveChanges:=True, FileName:=var_nombre_archivo
    oExcel.Quit
Set oWorkBook = Nothing
Set oSheet = Nothing
Set oExcel = Nothing
         
         
         MsgBox "Se a terminado de guardar el archivo " + archivo
      End If
   Else
      MsgBox "No se a seleccionado un archivo", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   MsgBox "A surgido un error al cargar el archivo, verifique que la hoja se llame PEDIDO y que las columnas sean CODIGO y PEDIDO", vbOKOnly, "ATENCION"
   If rs.State = 1 Then
      rs.Close
   End If
   If rsaux.State = 1 Then
      rsaux.Close
   End If
   If rsaux2.State = 1 Then
      rsaux2.Close
   End If
   If rsaux3.State = 1 Then
      rsaux3.Close
   End If
   If rsaux4.State = 1 Then
      rsaux4.Close
   End If
   If rsaux5.State = 1 Then
      rsaux5.Close
   End If
   If rsaux6.State = 1 Then
      rsaux6.Close
   End If
   If rsaux7.State = 1 Then
      rsaux7.Close
   End If
   If rsaux8.State = 1 Then
      rsaux8.Close
   End If
End Sub

Sub Ordena()
'
' Ordena Macro
' Macro grabada el 23/01/2010 por hlopez
'

'
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Sort Key1:=ActiveCell.Offset(-1, 0).Range("A1"), Order1:= _
        xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, _
        Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers
    ActiveCell.Offset(0, 5).Range("A1").Activate
    Selection.Sort Key1:=ActiveCell.Offset(-1, 0).Range("A1"), Order1:= _
        xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, _
        Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers
    ActiveCell.Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub



Private Sub cmd_imprimir_Click()
   If IsNumeric(Me.txt_pedido) Then
      Set reporte = appl.OpenReport(App.Path + "\rep_pedido_Cantia.rpt")
      frmvistasprevias.cr.ReportSource = reporte
      reporte.RecordSelectionFormula = "{VW_PEDIDO_UBICACIONES.INTE_PED_NUMERO} = " + Me.txt_pedido
      For ntablas = 1 To reporte.Database.Tables.Count
          reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_crystal, parametros(4), parametros(5)
      Next ntablas
      frmvistasprevias.cr.ViewReport
      frmvistasprevias.Caption = "Reporte de Movimientos"
      frmvistasprevias.Show 1
      Set reporte = Nothing
      var_si = MsgBox("Desea exportar el archivo", vbYesNo, "ATENCION")
      If var_si = 6 Then
         Set reporte = appl.OpenReport(App.Path + "\rep_pedido_Cantia_excel.rpt")
         reporte.RecordSelectionFormula = "{VW_PEDIDO_UBICACIONES.INTE_PED_NUMERO} = " + Me.txt_pedido
         For ntablas = 1 To reporte.Database.Tables.Count
             reporte.Database.Tables(ntablas).SetLogOnInfo parametros(6), var_bd_crystal, parametros(4), parametros(5)
         Next ntablas
         reporte.ExportOptions.FormatType = crEFTExcel80
         reporte.ExportOptions.DestinationType = crEDTDiskFile
         archivo = "c:\reporte_pedido_cantia_" & Replace(Str(Date), "/", "") & "_" & CStr(Hour(Time)) + "_" + CStr(Minute(Time)) + "_" + CStr(Second(Time)) + "_.xls"
         reporte.ExportOptions.DiskFileName = archivo
         reporte.Export False
         Set reporte = Nothing
         MsgBox "Se a terminado de guardar el archivo " + archivo
      End If
   Else
      MsgBox "Número de pedido incorrecto", vbOKOnly, "ATENCION"
   End If
End Sub

Private Sub Dir1_Change()
   Me.File1.Path = Me.Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error GoTo SALIR:
   Me.Dir1.Path = Me.Drive1.Drive
   Me.Dir1.Refresh
   Exit Sub
SALIR:
   MsgBox "Unidad incorrecta"
   Me.Drive1.Drive = "c:"
End Sub

Private Sub File1_Click()
   If CStr(Me.Dir1.Path) = "C:\" Or CStr(Me.Dir1.Path) = "c:\" Then
      Me.txt_ruta = CStr(Me.Dir1.Path) + Me.File1.FileName
   Else
      Me.txt_ruta = CStr(Me.Dir1.Path) + "\" + Me.File1.FileName
   End If
End Sub

Private Sub Form_Load()
   Top = 1500
   Left = 2300
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_articulos2)
End Sub

Private Sub txt_pedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.cmd_imprimir.SetFocus
   End If
End Sub


