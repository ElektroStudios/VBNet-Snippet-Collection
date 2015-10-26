# RegEx snippets category
These snippets are intended to help manaing Regex aspects.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editorcontext menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# RegEx Util
Contains related RegEx utilities

Public Members Summary

 - Constants
   - RegExUtil.Patterns.AlphabeticText As String
   - RegExUtil.Patterns.AlphanumericText As String
   - RegExUtil.Patterns.CreditCard As String
   - RegExUtil.Patterns.EMail As String
   - RegExUtil.Patterns.Hex As String
   - RegExUtil.Patterns.HtmlTag As String
   - RegExUtil.Patterns.Ipv4 As String
   - RegExUtil.Patterns.Ipv6 As String
   - RegExUtil.Patterns.NumericText As String
   - RegExUtil.Patterns.Phone As String
   - RegExUtil.Patterns.SafeText As String
   - RegExUtil.Patterns.Url As String
   - RegExUtil.Patterns.USphone As String
   - RegExUtil.Patterns.USssn As String
   - RegExUtil.Patterns.USstate As String
   - RegExUtil.Patterns.USzip As String

 - Child Classes
   - RegExUtil.Patterns

 - Types
   - RegExUtil.MatchPositionInfo <Serializable>

 - Constructors
   - RegExUtil.MatchPositionInfo.New(String, Integer)

 - Properties
   - RegExUtil.MatchPositionInfo.Text As String
   - RegExUtil.MatchPositionInfo.StartIndex As Integer
   - RegExUtil.MatchPositionInfo.EndIndex As Integer
   - RegExUtil.MatchPositionInfo.Length As Integer

 - Functions
   - RegExUtil.GetMatchesPositions(Regex, String, Integer) As IEnumerable(Of RegExUtil.MatchPositionInfo) 
   - RegExUtil.Validate(String, Boolean) As Boolean
