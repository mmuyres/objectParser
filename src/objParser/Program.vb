Namespace ObjParserCli

    Public Module Program

        Public Function Main(args As String()) As Integer
            Return Cli.Run(args, Console.In, Console.Out, Console.Error)
        End Function

    End Module

End Namespace
