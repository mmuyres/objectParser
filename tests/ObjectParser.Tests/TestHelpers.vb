Imports System.IO
Imports ObjectParser

Friend Module TestHelpers

    Public Function DataPath(name As String) As String
        Return Path.Combine(AppContext.BaseDirectory, "TestData", name)
    End Function

    Public Function ReadData(name As String) As String
        Return File.ReadAllText(DataPath(name))
    End Function

    ''' <summary>Runs <paramref name="path"/> over <paramref name="text"/> with a built-in map.</summary>
    Public Function Q(text As String, path As String, Optional mapName As String = "xml") As List(Of String)
        Dim parser = New Parser(TokenMap.BuiltIn(mapName), text)
        Return parser.Query(path).ToList()
    End Function

    ''' <summary>Runs <paramref name="path"/> over <paramref name="text"/> with a map given as text.</summary>
    Public Function QMap(mapText As String, text As String, path As String) As List(Of String)
        Dim parser = New Parser(TokenMap.Parse(mapText), text)
        Return parser.Query(path).ToList()
    End Function

    ''' <summary>Normalises line endings so assertions do not depend on how files were checked out.</summary>
    Public Function Lf(s As String) As String
        Return s.Replace(vbCrLf, vbLf)
    End Function

    Public Function LfAll(items As IEnumerable(Of String)) As List(Of String)
        Return items.Select(AddressOf Lf).ToList()
    End Function

End Module
