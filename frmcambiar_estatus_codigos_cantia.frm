VERSION 5.00
Begin VB.Form frmcambiar_estatus_codigos_cantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar estatus a artículos en cantia"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmbusqueda_pedido 
      Height          =   3900
      Left            =   75
      TabIndex        =   0
      Top             =   -15
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
         Caption         =   "Cambiar estatus"
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
         Caption         =   " Busqueda de pedido"
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
Attribute VB_Name = "frmcambiar_estatus_codigos_cantia"
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
      rsaux2.Open "SELECT * FROM [codigos$]", strConnectionString
      var_cadena_estatus = ""
      If Not rsaux2.EOF Then
         While Not rsaux2.EOF
               var_estatus = IIf(IsNull(rsaux2!Estatus), "0", rsaux2!Estatus)
               rsaux.Open "select * from tb_clasearticulos where vcha_car_clase_id = '" + CStr(var_estatus) + "'", cnn, adOpenDynamic, adLockOptimistic
               If rsaux.EOF Then
                  var_cadena_estatus = "El archivo contiene códigos con estatus inexistentes"
               End If
               rsaux.Close
               rsaux2.MoveNext
         Wend
         If var_cadena_estatus = "" Then
            rsaux2.MoveFirst
            While Not rsaux2.EOF
                  var_estatus = IIf(IsNull(rsaux2!Estatus), "0", rsaux2!Estatus)
                  rsaux.Open "update tb_Articulos set vcha_car_clase_id = '" + CStr(var_estatus) + "' where vcha_Art_articulo_id = '" + IIf(IsNull(rsaux2!codigo), "", rsaux2!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
                  rsaux2.MoveNext
            Wend
            MsgBox "Se a terminado el cambio de estatus", vbOKOnly, "ATENCION"
         Else
            MsgBox var_cadena_estatus, vbOKOnly, "ATENCION"
         End If
      Else
         MsgBox "El archivo no tiene información", vbOKOnly, "ATENCION"
      End If
      rsaux2.Close
   Else
      MsgBox "No se a seleccionado un archivo", vbOKOnly, "ATENCION"
   End If
   Exit Sub
SALIR:
   MsgBox "A surgido un error al cargar el archivo, puede que este no tenga el formato adecuado", vbOKOnly, "ATENCION"
   If rsaux2.State = 1 Then
      rsaux2.Close
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



