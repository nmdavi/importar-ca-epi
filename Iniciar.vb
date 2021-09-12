Imports System.IO
Imports System.Text
Imports ImportarCaEpi.Classes

Module Iniciar
    Private Log As New StringBuilder()
    Private PastaArquivo = ""

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
    ''' Baixar lista de CA's do site do Ministério do Trabalho
    ''' </summary>
    Private Sub ImportarCA()
        Dim importador As New FTP()

        'Log
        Log.Clear()
        Log.AppendLine("Iniciando o serviço")

        'FTP
        Try
            Log.AppendLine("Conectando ao FTP")

            importador.ConectarFTP("ftp://ftp.mtps.gov.br/portal/fiscalizacao/seguranca-e-saude-no-trabalho/caepi/tgg_export_caepi.txt")
            importador.Download()

            Log.AppendLine("Baixando arquivo ")

            importador.BaixarArquivo(PastaArquivo)

            Log.AppendLine("Baixado com sucesso ")
        Catch ex As Exception
            Log.AppendLine("Erro no FTP: " & ex.Message & " " & Trim(ex.StackTrace) & "")
            Return
        End Try
    End Sub

    ''' <summary>
    ''' Configurar pasta de destino dos arquivos
    ''' </summary>
    Private Sub InicializarValores()
        Try
            If Not Directory.Exists("C:\CaEpi") Then
                Directory.CreateDirectory("C:\CaEpi")
            End If

            'Para qual pasta irá o arquivo do Ministério do Trabalho
            Dim str As New StreamReader("Pasta.txt")
            PastaArquivo = str.ReadLine().Trim() 'Exemplo: C:\tgg_export_caepi.txt
            str.Close()
        Catch ex As Exception
            Log.AppendLine("Erro no InicializarValores" & ex.Message & " " & Trim(ex.StackTrace))
        End Try
    End Sub

    Private Sub GravarLog(ByVal msg As String)
        Dim stw As New StreamWriter("./Log.txt", True, Encoding.UTF8)
        stw.WriteLine(msg & " - " & Date.Now.ToString("dd/MM/yyyy HH:mm:ss"))
        stw.Close()
    End Sub
End Module
