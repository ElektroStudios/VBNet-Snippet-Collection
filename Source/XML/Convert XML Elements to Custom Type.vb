Public NotInheritable Class MyType

    Public ReadOnly Property MyProperty As String
        Get
            Return Me.myPropertyB
        End Get
    End Property
    Private ReadOnly myPropertyB As String

    Sub New(ByVal myParameter As String)
        MyClass.myPropertyB = myParameter
    End Sub

End Class

Private Sub Test()

    ' Fast way to load an XML document:
    ' Dim xml As XDocument = XDocument.Load(xmlfile) 

    Dim xml As XDocument =
        <?xml version="1.0" encoding="Windows-1252"?>
        <!--XML Songs Database-->
        <Songs>
            <Song><Name>My Song 1.mp3</Name></Song>
            <Song><Name>My Song 2.ogg</Name></Song>
            <Song><Name>My Song 3.wav</Name></Song>
        </Songs>

    Dim songList As IEnumerable(Of MyType) =
        From element As XElement In xml.<Songs>.<Song>
        Select New MyType(element.<Name>.Value)

    Dim sb As New StringBuilder
    For Each song As MyType In songList
        sb.AppendLine(String.Format("Name: {0}", song.MyProperty))
    Next song
    Console.WriteLine(sb.ToString)

End Sub
