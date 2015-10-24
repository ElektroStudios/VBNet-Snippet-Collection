# Xml snippets category
These snippets are intended to help managing Xml files or Xml operations.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editorcontext menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# Convert Xml Elements to Custom Type
A simple example for beginners, to convert Xml elements to custom Type.

# Xml Extensions
Contains custom extension methods to use with some of the "System.Xml" namespace members.
This adds additional functionalities to XDocument, XmlDocument and IEnumerable(Of XElement), for converting between types, sorting elements, getting x-paths, and more.

Public Members Summary
 - IEnumerable(Of XElement).DistinctByElement(String) As IEnumerable(Of XElement)
 - IEnumerable(Of XElement).SortByElement(String) As IEnumerable(Of XElement)
 - XDocument.DistinctByElement(String, String) As XDocument
 - XDocument.GetXPaths() As IEnumerable(Of String)
 - XDocument.SortByElement(String, String) As XDocument
 - XDocument.ToXmlDocument() As XmlDocument
 - XmlDocument.DistinctByElement(String, String) As XmlDocument
 - XmlDocument.GetXPaths() As IEnumerable(Of String)
 - XmlDocument.SortByElement(String, String) As XmlDocument
 - XmlDocument.ToXDocument() As XDocument

# Xml Util
Will contains related Xml utilities, such a Xml beautifier.
It also wraps an a "XmlTextWriter" to write an Xml document in a easier way for beginners.

Public Members Summary
 - XmlUtil.XmlBeautifier(String) As String
 - XmlUtil.EasyXmlTextWriter.New(String, Encoding)
 - XmlUtil.EasyXmlTextWriter.WriteBeginningElement(String, Formatting)
 - XmlUtil.EasyXmlTextWriter.WriteCommentLine(String, Formatting)
 - XmlUtil.EasyXmlTextWriter.WriteElement(String, String, Formatting)
 - XmlUtil.EasyXmlTextWriter.WriteElements(String(,), Formatting)
 - XmlUtil.EasyXmlTextWriter.WriteEndElement(Formatting)
 - XmlUtil.EasyXmlTextWriter.WriteEof(Boolean)
 - XmlUtil.EasyXmlTextWriter.WriteHeader()
