# Textfile snippets category
These snippets are intended to help managing plain text files.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editorcontext menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# TextFieldParser Example
Example to read a CSV file using TextFieldParser Class

# Textfile Stream
Reads and manages the contents of a textfile.
It encapsulates an underliying "FileStream" to access the file.

Public Members Summary

[+] Child Classes
 - TextfileStream.TexfileLines : Inherits List(Of String)

[+] Properties
 - TextfileStream.Filepath As String
 - TextfileStream.Encoding As Encoding
 - TextfileStream.Lines As TexfileLines
 - TextfileStream.Fs As FileStream
 - TextfileStream.FileHandle As Win32.SafeHandles.SafeFileHandle
 - 
 - TextfileStream.TexfileLines.CountBlank() As Integer
 - TextfileStream.TexfileLines.CountNonBlank() As Integer

[+] Functions
 - TextfileStream.ToString() As String

[+] Methods

 - TextfileStream.Lock()
 - TextfileStream.Unlock()
 - TextfileStream.Close()
 - TextfileStream.Dispose()
 - TextfileStream.Save(Opt: Encoding)
 - TextfileStream.Save(String, Encoding)
 - 
 - TextfileStream.TexfileLines.Randomize() As IEnumerable(Of String)
 - TextfileStream.TexfileLines.RemoveAt(IEnumerable(Of Integer)) As IEnumerable(Of String)
 - TextfileStream.TexfileLines.Trim(Opt: Char()) As IEnumerable(Of String)
 - TextfileStream.TexfileLines.TrimStart(Opt: Char()) As IEnumerable(Of String)
 - TextfileStream.TexfileLines.TrimEnd(Opt: Char()) As IEnumerable(Of String)
