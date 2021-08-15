Imports System.IO
Imports System.Net

Namespace Classes
    Public Class FTP
        Dim ftp As FtpWebRequest

        Public Sub ConectarFTP(ByVal url As String)
            ftp = WebRequest.Create(url)
        End Sub

        Public Sub Download()
            ftp.Method = WebRequestMethods.Ftp.DownloadFile
        End Sub

        Public Sub BaixarArquivo(ByVal pastaZip As String)
            Dim resposta As WebResponse = ftp.GetResponse()
            Dim gravador As Stream = resposta.GetResponseStream()
            Dim zip As Stream = File.Create(pastaZip)

            gravador.CopyTo(zip)

            gravador.Close()
            zip.Close()
            resposta.Close()
        End Sub

    End Class
End Namespace
