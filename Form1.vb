Imports System.IO
Imports System
Imports System.Threading
Public Class Form1
    Dim Cuenta As Integer
    Dim Mensaje As String
    Dim Ima2 As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Sfile As New StreamReader(Txtcorreos.Text, True)
        Dim SfileMSG As New StreamReader(TxtMsg.Text)
        Dim l As String

        While Not SfileMSG.EndOfStream
            l = SfileMSG.ReadLine()
            Mensaje = Mensaje & l
        End While
        SfileMSG.Close()

        While Not Sfile.EndOfStream
            l = Sfile.ReadLine()
            Label4.Text = l
            If Trim(l) <> "" And InStr(l, "@") > 0 Then correo(l)
            Thread.Sleep(1500)
        End While
        Sfile.Close()
        MsgBox("Terminado")
    End Sub
    Sub correo(ByVal para As String)
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient("smtp01.cmoderna.com", 26)
        _SMTP.Credentials = New System.Net.NetworkCredential("ecacerest", "c4c3r1t0s", "cmoderna")
        Dim adjunto As System.Net.Mail.Attachment
        'Dim adjunto2 As System.Net.Mail.Attachment
        adjunto = New System.Net.Mail.Attachment(TxtIma.Text)
        'adjunto2 = New System.Net.Mail.Attachment(Ima2)

        Try
            'CONFIGURACIÓN DEL STMP
            '_SMTP.Credentials = New System.Net.NetworkCredential("atorres@finagil.com.mx", "")
            '_SMTP.Host = "smtp01.cmoderna.com"
            '_SMTP.Port = 26
            '_SMTP.EnableSsl = True

            ' CONFIGURACION DEL MENSAJE
            'MessageBox.Show(para)
            _Message.[To].Add(para) 'Cuenta de Correo al que se le quiere enviar el e-mail
            _Message.From = New System.Net.Mail.MailAddress("gbello@finagil.com.mx", "Gabriel Bello (Finagil)", System.Text.Encoding.UTF8) 'Quien lo envía

            _Message.Subject = "COMUNICADO USO FRAUDULENTO DE INFORMACION DE FINAGIL"
            _Message.Priority = Net.Mail.MailPriority.High
            _Message.Body = Mensaje
            _Message.Priority = System.Net.Mail.MailPriority.High
            _Message.IsBodyHtml = True
            _Message.Attachments.Add(adjunto)
            '_Message.Attachments.Add(adjunto2)
            'ENVIO
            Dim ff As New StreamWriter("corr.txt", True)
            ff.WriteLine(para)
            ff.Close()

            _SMTP.Send(_Message)
            'Console.WriteLine("Mensaje enviado correctamene")
        Catch ex As System.Net.Mail.SmtpException
            'MsgBox(ex.ToString)
            Dim ff As New StreamWriter("corrERROR.txt", True)
            ff.WriteLine(para)
            ff.Close()
        End Try





    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TxtMsg.Text = Application.StartupPath & "\mensaje.htm"
        TxtIma.Text = Application.StartupPath & "\COMUNICADO.pdf"
        'Ima2 = Application.StartupPath & "\CARTA.docx"
        Txtcorreos.Text = Application.StartupPath & "\correos.txt"
    End Sub

End Class
