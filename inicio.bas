Attribute VB_Name = "inicio"
Dim var_tabla As ADODB.Connection
Dim var_fecha_local As Date
Global var_fecha_servidor As Date
Global var_archivo_local As String
Global var_archivo_servidor As String
Dim rs As ADODB.Recordset
Dim var_ruta As String
Dim var_ruta_local As String
Global parametros(9) As String
Global cnn As ADODB.Connection




Public Sub Main()
   'On Error GoTo salir:
   var_ruta_local = App.Path + "\"
   Dim var_n As Integer
   Dim var_i As Integer
   Set cnn = CreateObject("ADODB.connection")
   Set rs = CreateObject("ADODB.recordset")
   
   If Dir(App.Path + "\SID.SID") <> "" Then
   
      Open (App.Path + "\SID.SID") For Input As #1
      i = 0
      Do While Not EOF(1)
         Line Input #1, linea
         parametros(i) = linea
         i = i + 1
      Loop
    
      Close #1
      var_conexion_string = "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=" + parametros(1) + ";Data Source=" & parametros(0)
      'var_conexion_string = "Provider=SQLOLEDB.1;Password=elia;Persist Security Info=True;User ID=sa;Initial Catalog=sidalmacenbkp;Data Source=SQLQUEZADA2"
      cnn.Open var_conexion_string
    
      
      rs.Open "SELECT VCHA_PRI_RUTA_SISTEMA FROM TB_PRINCIPAL", cnn, adOpenDynamic, adLockOptimistic
      If Not rs.EOF Then
         var_ruta = IIf(IsNull(rs(0).Value), "", rs(0).Value)
      Else
         var_ruta = ""
         Open (App.Path + "\SID.SID") For Input As #1
         i = 0
         Do While Not EOF(1)
            Line Input #1, linea
            If i = 7 Then
               var_ruta = linea
            End If
            i = i + 1
         Loop
         Close #1
      End If
      rs.Close
   
   
   
   
      'Open (App.Path + "\SID.SID") For Input As #1
      'i = 0
      'Do While Not EOF(1)
      '   Line Input #1, linea
      '   If i = 7 Then
      '      var_ruta = linea
      '   End If
      '   i = i + 1
      'Loop
      'Close #1
      
      
      
      If IsNull(App.Path) = False Then
         If var_ruta <> "" Then
            var_archivo_local = Dir(Trim(var_ruta_local) + "SISTEMA.EXE")
            var_archivo_servidor = Dir(Trim(var_ruta) + "SISTEMA.EXE")
            If var_archivo_local = "sistema.exe" Then
               If var_archivo_servidor = "sistema.exe" Then
                  var_archivo_local = (Trim(var_ruta_local) + "SISTEMA.EXE")
                  var_fecha_local = FileDateTime(var_archivo_local)
                  var_archivo_servidor = (Trim(var_ruta) + "SISTEMA.EXE")
                  var_fecha_servidor = FileDateTime(var_archivo_servidor)
                  var_archivo_servidor = (Trim(var_ruta) + "SISTEMA.EXE")
                  If var_fecha_local < var_fecha_servidor Then
                     frmaviso_actualiza_sistema.Show 1
                  End If
                  frmejecuta.File1.Path = var_ruta
                  var_n = frmejecuta.File1.ListCount
                  For var_i = 1 To var_n
                      frmejecuta.File1.ListIndex = var_i - 1
                      var_reporte = frmejecuta.File1.FileName
                      var_archivo_local = Dir(Trim(var_ruta_local) + Trim(var_reporte))
                      var_archivo_servidor = Dir(Trim(var_ruta) + Trim(var_reporte))
                      If var_archivo_local = Trim(var_reporte) Then
                         If var_archivo_servidor = Trim(var_reporte) Then
                            var_archivo_local = (Trim(var_ruta_local) + Trim(var_reporte))
                            var_fecha_local = FileDateTime(var_archivo_local)
                            var_archivo_servidor = (Trim(var_ruta) + Trim(var_reporte))
                            var_fecha_servidor = FileDateTime(var_archivo_servidor)
                            var_archivo_servidor = (Trim(var_ruta) + Trim(var_reporte))
                            If var_fecha_local < var_fecha_servidor Then
                               FileCopy var_archivo_servidor, var_archivo_local
                            End If
                         End If
                      Else
                         var_archivo_local = (Trim(var_ruta_local) + Trim(var_reporte))
                         var_archivo_servidor = (Trim(var_ruta) + Trim(var_reporte))
                         FileCopy var_archivo_servidor, var_ruta_local + Trim(var_reporte)
                      End If
                  Next var_i
                  var_archivo_local = Trim(var_ruta_local) + "SISTEMA.EXE"
                  x = Shell(var_archivo_local, vbNormalFocus)
                  End
               Else
                  MsgBox "No existe una actualización del sistema", vbOKOnly, "ATENCION"
                  End
               End If
            Else
               x = Shell(App.Path + "sistema.exe", vbNormalFocus)
               End
            End If
         Else
            MsgBox "No existe una ruta para la actualización del SID", vbOKOnly, "ATENCION"
         End If
      End If
   Else
      MsgBox "No existe el archivo SID.SID, favor de verificarlo con el administrador del sistema", vbOKOnly, "ATENCION"
      End
   End If
   Exit Sub
salir:
   If Err.Number = 52 Then
      MsgBox "Es posible que no este conectado a la red o que no tenga acceso a la carpeta " + var_ruta, vbOKOnly, "ATENCION"
   Else
      MsgBox "Existe un problema en la configuración del sistema, favor de verificarlo con el administrador del sistema", vbOKOnly, "ATENCION"
   End If
   End
End Sub
