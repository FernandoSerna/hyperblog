Imports System.Net.Sockets
Imports System.Net
Imports System.IO.Ports
Imports System.Text
Imports textinsdk
Imports System.Reflection


Public Class frmenvia_texto
    Public var_arreglo As String = Command()

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim v_parametros() As String = Split(var_arreglo, ",")
        Dim v_ip As String = v_parametros(0)
        Dim v_puerto As String = v_parametros(1)
        Dim v_texto As String = v_parametros(2)
        On Error GoTo salir
        txtIn.tcpIpSendData(v_ip, v_puerto, "@B@ " & v_texto & " @E@" & vbCrLf)
        Me.Close()
        Exit Sub
salir:
        Resume
    End Sub
End Class
