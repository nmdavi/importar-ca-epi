Imports System.IO
Imports System.Text
Imports ImportarCaEpi.Classes
Imports SharpCompress.Archives
Imports SharpCompress.Archives.Rar
Imports SharpCompress.Common

Module Iniciar
    Private Log As New StringBuilder()
    Private PastaZip = ""
    Private PastaDestino = ""

    Sub Main()
        '1º
        InicializarValores()
        '2º
        ImportarCA()

        GravarLog(Log.ToString)

        Console.WriteLine(Log.ToString())
        Console.WriteLine("O arquivo se encontra na pasta C:\CaEpi")
        Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Importar e extrair lista de CA's do site do Ministério do Trabalho
    ''' </summary>
    Private Sub ImportarCA()
        Dim importador As New FTP()

        'Log
        Log.Clear()
        Log.AppendLine("Iniciando o serviço")

        'FTP
        Try
            Log.AppendLine("Conectando ao FTP")

            importador.ConectarFTP("ftp://ftp.mtps.gov.br/portal/fiscalizacao/seguranca-e-saude-no-trabalho/caepi/tgg_export_caepi.zip")
            importador.Download()

            Log.AppendLine("Baixando arquivo ")

            importador.BaixarArquivo(PastaZip)

            Log.AppendLine("Baixado com sucesso ")
        Catch ex As Exception
            Log.AppendLine("Erro no FTP: " & ex.Message & " " & Trim(ex.StackTrace) & "")
            Return
        End Try

        'Descompactar com SharpCompress
        Try
            Log.AppendLine("Descompactando com SharpCompress")

            Dim archive As RarArchive = RarArchive.Open(PastaZip)
            For Each entry As RarArchiveEntry In archive.Entries.Where(Function(x) Not x.IsDirectory)
                entry.WriteToDirectory(PastaDestino, New ExtractionOptions() With {.ExtractFullPath = True, .Overwrite = True})
            Next
            'archive.Dispose()

            Log.AppendLine("Descompactado com sucesso")

            'Não executo o UnRar quando o SharpCompress funcionar
            Return
        Catch ex As Exception
            'Se der erro tento pelo UnRar
            Log.AppendLine("Erro no SharpCompress: " & ex.Message & " " & Trim(ex.StackTrace) & "")
        End Try

        'Descompactar com UnRar (mesmos criadores do WinRar)
        Try
            Log.AppendLine("Descompactando com UnRar")

            Dim unRar As New Decompressor(PastaZip)
            unRar.UnPackAll(PastaDestino)

            Log.AppendLine("Descompactado com sucesso")
        Catch ex As Exception
            'Se der erro gravo no log e encerro a execução
            Log.AppendLine("Erro no UnRar: " & ex.Message & " " & Trim(ex.StackTrace) & "")
        End Try
    End Sub

    ''' <summary>
    ''' Configurar pasta de destino dos arquivos
    ''' </summary>
    Private Sub InicializarValores()
        Dim pasta As String()

        Try
            If Not Directory.Exists("C:\CaEpi") Then
                Directory.CreateDirectory("C:\CaEpi")
            End If

            'Pasta
            Dim str As New StreamReader("Pasta.txt")
            pasta = str.ReadLine().Split("|")
            str.Close()

            'Para qual pasta irá o ZIP do Ministério do Trabalho
            PastaZip = pasta(0) ' Exemplo: C:\tgg_export_caepi.zip

            'Para qual pasta irá os arquivos do ZIP do Ministério do Trabalho
            PastaDestino = pasta(1) 'Exemplo: C:\

        Catch ex As Exception
            Log.AppendLine("Erro no InicializarValores" & ex.Message & " " & Trim(ex.StackTrace) & "")
        End Try
    End Sub

    Private Sub GravarLog(ByVal msg As String)
        Dim stw As New StreamWriter("./Log.txt", True, Encoding.UTF8)
        stw.WriteLine(msg & " - " & Date.Now.ToString("dd/MM/yyyy HH:mm:ss"))
        stw.Close()
    End Sub
End Module
