VERSION 5.00
Begin VB.Form frmimportar_informacion_articulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de Información de Artículos"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_importar 
      Caption         =   "Importación de Información de Artículos"
      Height          =   645
      Left            =   75
      TabIndex        =   0
      Top             =   270
      Width           =   4530
   End
End
Attribute VB_Name = "frmimportar_informacion_articulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var_tabla As ADODB.Connection
Private Sub Command1_Click()
End Sub

Private Sub cmd_importar_Click()
   Dim var_volumen As Double
   If var_tabla.State = 1 Then
      var_tabla.Close
   End If
   var_tabla.Open "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=" + """" + "MSDASQL.1;Persist Security Info=False;DSN=Visual FoxPro Tables;UID=;SourceDB=" + App.Path + ";SourceType=DBF;Exclusive=No;BackgroundFetch=Sí;Collate=Machine;" + """"
   rs.Open "select * from volumenes", var_tabla, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
         var_volumen = IIf(IsNull(rs!volumen), 0, rs!volumen)
         rsaux.Open "select * from tb_equivalencias where vcha_equ_codigo_equivalente = '" + Trim(rs!codigo) + "'", cnn, adOpenDynamic, adLockOptimistic
         If Not rsaux.EOF Then
            var_codigo = IIf(IsNull(rsaux!vcha_Art_Articulo_id), "", rsaux!vcha_Art_Articulo_id)
            rsaux2.Open "update tb_Articulos set floa_art_volumen = " + CStr(var_volumen) + " where vcha_art_articulo_id = '" + rsaux!vcha_Art_Articulo_id + "'", cnn, adOpenDynamic, adLockOptimistic
         End If
         rsaux.Close
         rs.MoveNext
   Wend
   rs.Close
End Sub

Private Sub Form_Load()
   Top = 3000
   Left = 3500
   Set var_tabla = CreateObject("ADODB.connection")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   activa_forma (var_activa_forma_articulos2)
End Sub
