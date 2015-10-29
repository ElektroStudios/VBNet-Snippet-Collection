# Error-Handling snippets category
These snippets are intended to help managing error-handling and logging scenarios.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# Logfile Writer
A (very)simple logging system assistant that helps to create a log for the current application.

Public Members Summary

 - Properties
   - LogfileWriter.Filepath As String
   - LogfileWriter.Encoding As Encoding
   - LogfileWriter.EntryFormat As String

 - Methods
   - LogfileWriter.WriteEntry(TraceEventType, String)
   - LogfileWriter.WriteNewLine()
   - LogfileWriter.Clear()
