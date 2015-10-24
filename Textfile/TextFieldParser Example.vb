' The CSV content to read.
Dim csvReader As New StringReader("@Username, Password, Privileges" &
                                  Environment.NewLine &
                                  "Elektro; ""My Password""; Administrator" &
                                  Environment.NewLine &
                                  "Guest; none; Administrator")

' The TextFieldParser instance.
Dim csvParser As New TextFieldParser(reader:=csvReader) With
    {
        .Delimiters = {";"},
        .CommentTokens = {"@"},
        .HasFieldsEnclosedInQuotes = True,
        .TextFieldType = FieldType.Delimited,
        .TrimWhiteSpace = True
    }

' Iterate the CSV lines
Do Until csvParser.EndOfData

    ' Current line.
    Dim csvLine As String = csvReader.ReadLine

    ' Current fields.
    Dim csvFields As IEnumerable(Of String) = csvParser.ReadFields()

    For Each field As String In csvFields
        Console.WriteLine(field)
    Next field

Loop
