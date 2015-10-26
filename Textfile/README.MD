# Textfile snippets category
These snippets are intended to help managing plain text files.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editorcontext menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# TextFieldParser Example
Example for beginners, to read a CSV file using "TextFieldParser" Class

# Textfile Stream
Reads and manages the contents of a textfile.
It encapsulates an underliying "FileStream" to access the file.

Public Members Summary

 - Child Classes
   - TextfileStream.TexfileLines : Inherits List(Of String)

 - Properties
   - TextfileStream.Filepath As String
   - TextfileStream.Encoding As Encoding
   - TextfileStream.Lines As TexfileLines
   - TextfileStream.Fs As FileStream
   - TextfileStream.FileHandle As Win32.SafeHandles.SafeFileHandle
   - TextfileStream.TexfileLines.CountBlank() As Integer
   - TextfileStream.TexfileLines.CountNonBlank() As Integer

 -  Functions
   - TextfileStream.ToString() As String

 -  Methods
   - TextfileStream.Lock()
   - TextfileStream.Unlock()
   - TextfileStream.Close()
   - TextfileStream.Dispose()
   - TextfileStream.Save(Opt: Encoding)
   - TextfileStream.Save(String, Encoding)
   - TextfileStream.TexfileLines.Randomize() As IEnumerable(Of String)
   - TextfileStream.TexfileLines.RemoveAt(IEnumerable(Of Integer)) As IEnumerable(Of String)
   - TextfileStream.TexfileLines.Trim(Opt: Char()) As IEnumerable(Of String)
   - TextfileStream.TexfileLines.TrimStart(Opt: Char()) As IEnumerable(Of String)
   - TextfileStream.TexfileLines.TrimEnd(Opt: Char()) As IEnumerable(Of String)

 -  Constructors
   -  TextfileStream.New(String, Opt: Encoding)

#Textfile Util
Contains related text-file utilities.

Public Members Summary

 - Functions
   - TextfileUtil.Contains(String, String, IEqualityComparer(Of String), Encoding) As Boolean
   - TextfileUtil.CountAllLines(String, Opt: Encoding) As Integer
   - TextfileUtil.CountBlankLines(String, Opt: Encoding) As Integer
   - TextfileUtil.CountNonBlankLines(String, Opt: Encoding) As Integer
   - TextfileUtil.GetEncoding(String) As Encoding
   - TextfileUtil.GetLine(String, Integer, TextfileUtil.TextDirection, Opt: Encoding) As String
   - TextfileUtil.GetLines(String, Integer(), TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.InsertLine(String, Integer, String(), TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.InsertLine(String, Integer, String, TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.Randomize(String, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.RemoveLine(String, Integer, TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.RemoveLines(String, Integer(), TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.ReplaceLine(String, Integer, String(), TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.ReplaceLine(String, Integer, String, TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.ReplaceLines(String, Integer(), String(), TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.ReplaceLines(String, Integer(), String, TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.Reverse(String, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.SkipLines(String, Integer, TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.Sort(String, ListSortDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.SortBy(Of T)(String, ListSortDirection, Func(Of String, T), Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.Split(String, Char(),  Opt: StringSplitOptions, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.Split(String, Char,  Opt: StringSplitOptions, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.Split(String, String(),  Opt: StringSplitOptions, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.Split(String, String,  Opt: StringSplitOptions, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.SplitByLines(String, Integer, Opt: Encoding) As IEnumerable(Of IEnumerable(Of String))
   - TextfileUtil.TakeLines(String, Integer, TextfileUtil.TextDirection, Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.Trim(String, Opt: Char(), Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.TrimEnd(String, Opt: Char(), Opt: Encoding) As IEnumerable(Of String)
   - TextfileUtil.TrimStart(String, Opt: Char(), Opt: Encoding) As IEnumerable(Of String)

 - Methods
   - TextfileUtil.InsertLine(String, String, Integer, String(), TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.InsertLine(String, String, Integer, String, TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.Randomize(String, String, Opt: Encoding)
   - TextfileUtil.RemoveLine(String, String, Integer, TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.RemoveLines(String, String, Integer(), TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.ReplaceLine(String, String, Integer, String(), TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.ReplaceLine(String, String, Integer, String, TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.ReplaceLines(String, String, Integer(), String(), TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.ReplaceLines(String, String, Integer(), String, TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.Reverse(String, String, Opt: Encoding)
   - TextfileUtil.SkipLines(String, String, Integer, TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.Sort(String, String, ListSortDirection, Opt: Encoding)
   - TextfileUtil.SortBy(Of T)(String, String, ListSortDirection, Func(Of String, T), Opt: Encoding)
   - TextfileUtil.TakeLines(String, String, Integer, TextfileUtil.TextDirection, Opt: Encoding)
   - TextfileUtil.Trim(String, String, Opt: Char(), Opt: Encoding)
   - TextfileUtil.TrimEnd(String, String, Opt: Char(), Opt: Encoding)
   - TextfileUtil.TrimStart(String, String, Opt: Char(), Opt: Encoding)
