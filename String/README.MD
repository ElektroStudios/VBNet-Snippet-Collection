# String snippets category
These snippets are intended to help managing Strings.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

#Multiline String Example
Simple example of a multiline String for beginners.

#String Extensions
Contains custom extension methods to use with a String datatype.

Measure a string, randomize a string, convert and/or detect a string case, get MD5 hash, 
convert to or from Hex or Bin, rotate string, expand string, replace or find by Regex, and much more.

Public Members Summary

 - Types
   - StringExtensions.GroupingInfo <Serializable>
   - StringExtensions.GroupingCharsInfo <Serializable>

 - Enumerations
   - StringExtensions.StringCase As Integer
   - StringExtensions.StringDirection As Integer

 - Properties
   - StringExtensions.GroupingCharsInfo.Aperture As Char
   - StringExtensions.GroupingCharsInfo.Closure As Char
   - StringExtensions.GroupingInfo.Source As String
   - StringExtensions.GroupingInfo.GroupingChars As List(Of GroupingCharsInfo)
   - StringExtensions.GroupingInfo.HasClosedGroups As Boolean
   - StringExtensions.GroupingInfo.HasOpenGroups As Boolean
   - StringExtensions.GroupingInfo.GroupsClosedCount As Integer
   - StringExtensions.GroupingInfo.GroupsOpenCount As Integer
   - StringExtensions.GroupingInfo.GroupsClosedPositions As Dictionary(Of Integer, Integer)
   - StringExtensions.GroupingInfo.GroupsOpenPositions As List(Of Integer)

 - Extension Functions
   - String.BinaryToAscII() As String
   - String.CountCharacter(Char) As Integer
   - String.CountCharacters(Char, StringComparison) As Integer
   - String.Delimit(String, Integer, Opt: RegexOptions) As String
   - String.Delimit(String, Opt: RegexOptions) As String
   - String.Delimit(String, String, Integer, Opt: RegexOptions) As String
   - String.Delimit(String, String, Opt: RegexOptions) As String
   - String.ExpandBlankSpace(Char, Opt: Integer) As String
   - String.ExpandBlankSpace(Opt: Integer) As String
   - String.Expandchars(Char, Opt: Integer) As String
   - String.Expandchars(Opt: Integer) As String
   - String.ExpandVariables() As String
   - String.GetGroupingInfo(Opt: List(Of StringExtensions.GroupingCharsInfo)) As StringExtensions.GroupingInfo
   - String.HasDiacritics() As Boolean
   - String.HexToAscII() As String
   - String.IndexOfAnyRegEx(IEnumerable(Of String), Integer, Integer, RegexOptions) As Integer
   - String.IndexOfAnyRegEx(IEnumerable(Of String), Integer, RegexOptions) As Integer
   - String.IndexOfAnyRegEx(IEnumerable(Of String), RegexOptions) As Integer
   - String.IndexOfRegEx(String, Integer, Integer, RegexOptions) As Integer
   - String.IndexOfRegEx(String, Integer, RegexOptions) As Integer
   - String.IndexOfRegEx(String, RegexOptions) As Integer
   - String.IsAlphabetic() As Boolean
   - String.IsCase(StringExtensions.StringCase) As Boolean
   - String.IsNumeric() As Boolean
   - String.IsNumericOf(Of T) As Boolean
   - String.IsValueOf(Of T) As Boolean
   - String.LastIndexOfAnyRegEx(IEnumerable(Of String), Integer, Integer, RegexOptions) As Integer
   - String.LastIndexOfAnyRegEx(IEnumerable(Of String), Integer, RegexOptions) As Integer
   - String.LastIndexOfAnyRegEx(IEnumerable(Of String), RegexOptions) As Integer
   - String.LastIndexOfRegEx(String, Integer, Integer, RegexOptions) As Integer
   - String.LastIndexOfRegEx(String, Integer, RegexOptions) As Integer
   - String.LastIndexOfRegEx(String, RegexOptions) As Integer
   - String.Measure(Font) As Size
   - String.NormalizeDiacritics(NormalizationForm) As String
   - String.Randomize() As String
   - String.Randomize(Integer) As String
   - String.Rename(StringExtensions.StringCase) As String
   - String.Replace(String, String, StringComparison) As String
   - String.ReplaceRegEx(String, String, RegExOptions) As String
   - String.Reverse() As String
   - String.Rotate(StringExtensions.StringDirection, Integer) As String
   - String.SizeInFile(Encoding) As Integer
   - String.SizeInFile(Encoding, Boolean) As Integer
   - String.SizeInMemory(Font) As Integer
   - String.SplitByLength(Integer) As IEnumerable(Of String)
   - String.SplitByPosition(Integer) As IEnumerable(Of String)
   - String.ToBase64() As String
   - String.ToBinary(Opt: String, Opt: Encoding) As String
   - String.ToByteArray(Opt: Encoding) As Byte()
   - String.ToHex(Opt: String, Opt: Encoding) As String
   - String.ToHtmlDocument() As Htmldocument
   - String.ToMD5() As String

 - Constructors
   - StringExtensions.GroupingCharsInfo.New(Char, Char)
   - StringExtensions.GroupingInfo.New(String, List(Of GroupingCharsInfo), Boolean, Boolean, Integer, Integer, Dictionary(Of Integer, Integer), List(Of Integer))

#String Util
Contains related String utilities, such a fixed-length string implementation.

Public Members Summary

 - Types
   - StringUtil.FixedLengthString <Serializable>

 - Properties
   - StringUtil.FixedLengthString.FixedLength As Integer
   - StringUtil.FixedLengthString.PaddingChar As Char
   - StringUtil.FixedLengthString.ValueUnfixed As String
   - StringUtil.FixedLengthString.ValueFixed As String

 - Functions
   - StringUtil.GetWhiteSpacedString(Integer) As String
   - StringUtil.FixedLengthString.ToString() As String

 - Constructors
   - StringUtil.FixedLengthString.New(String, Integer, Opt: Char)
