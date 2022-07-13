Imports Microsoft.Office.Interop


Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Frm_Main_Load(ByVal sender As Global.System.Object, ByVal e As Global.System.EventArgs) Handles MyBase.Load
        Dim var_parametro_1 As String
        Dim var_parametro_2 As String
        Dim var_parametro_3 As String
        Dim m_OutLook As Outlook.Application
        'Puedes verificar si te estan pasando los parametros
        If My.Application.CommandLineArgs.Count > 0 Then
            'Y asi los capturas, en unas variables tipo string para mi caso
            'La solucion esta planteada, el resto es tuyo
            var_parametro_1 = My.Application.CommandLineArgs(0)
            var_parametro_2 = My.Application.CommandLineArgs(1)
            var_parametro_3 = My.Application.CommandLineArgs(2)
            MsgBox(var_parametro_1)
            MsgBox(var_parametro_2)
            MsgBox(var_parametro_3)
            Try
                '* Creamos un Objeto tipo Mail 
                Dim objMail As Outlook.MailItem
                '* Inicializamos nuestra apliación OutLook 
                m_OutLook = New Outlook.Application
                '* Creamos una instancia de un objeto tipo MailItem 
                objMail = m_OutLook.CreateItem(Outlook.OlItemType.olMailItem)
                '* Asignamos las propiedades a nuestra Instancial del objeto 
                '* MailItem 
                objMail.To = "fserna@vianney.com.mx"
                objMail.Subject = "Enviando correo desde VB2010 .NET"
                objMail.Body = "El fer me cae bien pero bien gordo"

                objMail.Attachments.Add("c:\sistemas\FAEERE476.FAC")
                'Dim Rta = MsgBox("¿Realmente desea enviar el correo?", MsgBoxStyle.YesNo)
                Dim rta = 6
                If Rta = 6 Then
                    '* Enviamos nuestro Mail y listo! 
                    objMail.Send()
                    '* Desplegamos un mensaje indicando que todo fue exitoso 
                    MessageBox.Show("Envìo exitoso.", "Enviar Mail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                ElseIf Rta = 7 Then
                    MessageBox.Show("Eío cancelado", "Enviar Mail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If

            Catch ex As Exception
                '* Si se produce algun Error 
                MsgBox(Err.Description)
                MessageBox.Show("Error enviando mail")
            Finally
                m_OutLook = Nothing ' Destruimos el objeto (recoger la basura...)
            End Try
            End



        End If
    End Sub

  

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim m_OutLook As Outlook.Application
        Try
            '* Creamos un Objeto tipo Mail 
            Dim objMail As Outlook.MailItem
            '* Inicializamos nuestra apliación OutLook 
            m_OutLook = New Outlook.Application
            '* Creamos una instancia de un objeto tipo MailItem 
            objMail = m_OutLook.CreateItem(Outlook.OlItemType.olMailItem)
            '* Asignamos las propiedades a nuestra Instancial del objeto 
            '* MailItem 
            objMail.To = "fserna@vianney.com.mx"
            objMail.Subject = "Enviando correo desde VB2010 .NET"
            objMail.Body = "El fer me cae bien pero bien gordo"

            objMail.Attachments.Add("c:\sistemas\FAEERE476.FAC")
            Dim Rta = MsgBox("¿Realmente desea enviar el correo?", MsgBoxStyle.YesNo)
            If Rta = 6 Then
                '* Enviamos nuestro Mail y listo! 
                objMail.Send()
                '* Desplegamos un mensaje indicando que todo fue exitoso 
                MessageBox.Show("Envìo exitoso.", "Enviar Mail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            ElseIf Rta = 7 Then
                MessageBox.Show("Eío cancelado", "Enviar Mail", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

        Catch ex As Exception
            '* Si se produce algun Error 
            MessageBox.Show("Error enviando mail")
        Finally
            m_OutLook = Nothing ' Destruimos el objeto (recoger la basura...)
        End Try
    End Sub
End Class
