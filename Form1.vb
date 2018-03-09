Imports System.IO
Imports System
Public Class Form1
    Dim Cuenta As Integer
    Dim Mensaje As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Sfile As New StreamReader(Txtcorreos.Text)
        Dim SfileMSG As New StreamReader(TxtMsg.Text)
        Dim l As String

        While Not SfileMSG.EndOfStream
            l = SfileMSG.ReadLine()
            Mensaje = Mensaje & l
        End While
        SfileMSG.Close()

        While Not Sfile.EndOfStream
            l = Sfile.ReadLine()
            If Trim(l) <> "" And InStr(l, "@") > 0 Then correo(l)
        End While
        Sfile.Close()
        MsgBox("Terminado")
    End Sub
    Sub correo(ByVal para As String)
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim adjunto As System.Net.Mail.Attachment
        adjunto = New System.Net.Mail.Attachment(TxtIma.Text)
        If Cuenta = 0 Then Cuenta = 1
        Try
            'CONFIGURACIÓN DEL STMP
            Select Case Cuenta
                Case 1
                    _SMTP.Credentials = New System.Net.NetworkCredential("pnltoluca@gmail.com", "19670305")
                Case 2
                    _SMTP.Credentials = New System.Net.NetworkCredential("Patriciamtzcastillo@gmail.com", "19670305")
                Case 3
                    _SMTP.Credentials = New System.Net.NetworkCredential("estrapersonales@gmail.com", "14101973")

            End Select


            _SMTP.Host = "smtp.gmail.com"
            _SMTP.Port = 587
            _SMTP.EnableSsl = True

            ' CONFIGURACION DEL MENSAJE
            _Message.[To].Add(para) 'Cuenta de Correo al que se le quiere enviar el e-mail
            _Message.From = New System.Net.Mail.MailAddress("patriciamtzcastillo@gmail.com", "Patricia Martinez", System.Text.Encoding.UTF8) 'Quien lo envía

            _Message.Subject = "Photoreading"
            _Message.Body = Mensaje
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = True
            _Message.Attachments.Add(adjunto)
            'ENVIO

            _SMTP.Send(_Message)
            'Console.WriteLine("Mensaje enviado correctamene")
        Catch ex As System.Net.Mail.SmtpException
            MsgBox(ex.ToString)
            If MessageBox.Show("Cambio de Cuenta", "Cuenta", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Select Case Cuenta
                    Case 1
                        Cuenta = 2
                    Case 2
                        Cuenta = 3
                    Case 3
                        Cuenta = 1
                End Select
            End If
        End Try





    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TxtMsg.Text = Application.StartupPath & "\mensaje.htm"
        TxtIma.Text = Application.StartupPath & "\image002.jpg"
        Txtcorreos.Text = Application.StartupPath & "\correos.txt"
    End Sub
End Class
