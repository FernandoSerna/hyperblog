VERSION 5.00
Begin VB.Form frmetiquetas_ubicaciones_contingencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_codigo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmetiquetas_ubicaciones_contingencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Len(Me.txt_codigo) > 8 Then
      Me.txt_codigo = Mid(Me.txt_codigo, 4, 5)
   End If
   If Len(Me.txt_codigo) = 8 Then
      Me.txt_codigo = Mid(Me.txt_codigo, 4, 5)
   End If
   If Len(Me.txt_codigo) = 5 Then
      Me.txt_codigo = "000" + Me.txt_codigo
      rs.Open "select * from xxvia_system_items_b where segment1 = '" + Me.txt_codigo + "' and organization_id = 93", cnnoracle_4, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         Set fs = CreateObject("Scripting.FileSystemObject")
         Set A = fs.CreateTextFile(App.Path + "\etiquetas.txt", True)
         A.writeline ("")
         A.writeline ("US")
         A.writeline ("N")
         A.writeline ("q816")
         A.writeline ("Q1015,20+0")
         A.writeline ("S2")
         A.writeline ("D8")
         A.writeline ("ZT")
         A.writeline ("TTh:m")
         A.writeline ("TDy2.mn.dd")
         var_longitud = Len(Trim(rs!Description))
         var_ubicacion = IIf(IsNull(rs!attribute2), "", rs!attribute2)
         If var_longitud >= 35 Then
            var_articulo = Replace(Left(Trim(rsaux3!vcha_Art_nombre_español), 35), """", " ") + "  "
         End If
         If var_longitud < 35 Then
            var_articulo = Replace(Trim(rs!Description), """", " ")
            While Not var_longitud = 38
                  var_articulo = var_articulo + " "
                  var_longitud = var_longitud + 1
            Wend
         End If
         A.writeline ("A740,200,1,5,2,1,N,""" + var_ubicacion + """")
         A.writeline ("A400,400,1,5,2,1,N,""" + Me.txt_codigo + """")
         A.writeline ("A200,50,1,4,2,1,N,""" + var_articulo + """")
         A.writeline ("P1")
         A.Close
         
         Open (App.Path & "\net_use.bat") For Output As #3
         var_archivo = App.Path & "\net_use.bat"
         Print #3, "net use lpt1 \\" + fun_NombrePc + "\zebra"
         Close #3
         x = Shell(var_archivo, vbHide)
               
               
         Open (App.Path & "\etiquetas.bat") For Output As #2
         var_archivo = App.Path & "\etiquetas.bat"
         'Print #2, "copy " + App.Path + "\etiquetas.txt lpt1"
         Print #2, "copy " + App.Path + "\etiquetas.txt \\" + fun_NombrePc + "\zebra"
                  
         Close #2
         x = Shell(var_archivo, vbHide)
         Me.txt_codigo.Text = ""
          
      End If
      rs.Close
   End If

End If
End Sub
