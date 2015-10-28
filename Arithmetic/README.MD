# Arithmetic snippets category
These snippets are intended to help managing Xml files or Xml operations.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# Numeric Extensions
Contains custom extension methods to use with the Short, UShort, Integer, UInteger, 
Long, ULong, Decimal, Single and Double data-types.

This module adds additional functionalities for most numeric types to easy get percentages or portions, 
determine if a value is divisible by or multiple of, if its positive, negative or prime,
if its in range of values, get the difference of value, convert to VB or C# Hexadecimal literal, and much more.

Public Members Summary
 - Extension Functions
   - Short.DifferenceOf(Double) As Double
   - UShort.DifferenceOf(Double) As Double
   - Integer.DifferenceOf(Double) As Double
   - UInteger.DifferenceOf(Double) As Double
   - Long.DifferenceOf(Double) As Double
   - ULong.DifferenceOf(Double) As Double
   - Decimal.DifferenceOf(Double) As Double
   - Single.DifferenceOf(Double) As Double
   - Double.DifferenceOf(Double) As Double
   - Short.IsDivisibleBy(Double) As Boolean
   - UShort.IsDivisibleBy(Double) As Boolean
   - Integer.IsDivisibleBy(Double) As Boolean
   - UInteger.IsDivisibleBy(Double) As Boolean
   - Long.IsDivisibleBy(Double) As Boolean
   - ULong.IsDivisibleBy(Double) As Boolean
   - Decimal.IsDivisibleBy(Double) As Boolean
   - Single.IsDivisibleBy(Double) As Boolean
   - Double.IsDivisibleBy(Double) As Boolean
   - Short.IsInRangeOf(Double, Double) As Boolean
   - UShort.IsInRangeOf(Double, Double) As Boolean
   - Integer.IsInRangeOf(Double, Double) As Boolean
   - UInteger.IsInRangeOf(Double, Double) As Boolean
   - Long.IsInRangeOf(Double, Double) As Boolean
   - ULong.IsInRangeOf(Double, Double) As Boolean
   - Decimal.IsInRangeOf(Double, Double) As Boolean
   - Single.IsInRangeOf(Double, Double) As Boolean
   - Double.IsInRangeOf(Double, Double) As Boolean
   - Short.IsPositive() As Boolean
   - UShort.IsPositive() As Boolean
   - Integer.IsPositive() As Boolean
   - UInteger.IsPositive() As Boolean
   - Long.IsPositive() As Boolean
   - ULong.IsPositive() As Boolean
   - Decimal.IsPositive() As Boolean
   - Single.IsPositive() As Boolean
   - Double.IsPositive() As Boolean
   - Short.IsNegative() As Boolean
   - UShort.IsNegative() As Boolean
   - Integer.IsNegative() As Boolean
   - UInteger.IsNegative() As Boolean
   - Long.IsNegative() As Boolean
   - ULong.IsNegative() As Boolean
   - Decimal.IsNegative() As Boolean
   - Single.IsNegative() As Boolean
   - Double.IsNegative() As Boolean
   - Short.IsMultipleBOf(Double) As Boolean
   - UShort.IsMultipleBOf(Double) As Boolean
   - Integer.IsMultipleBOf(Double) As Boolean
   - UInteger.IsMultipleBOf(Double) As Boolean
   - Long.IsMultipleBOf(Double) As Boolean
   - ULong.IsMultipleBOf(Double) As Boolean
   - Decimal.IsMultipleBOf(Double) As Boolean
   - Single.IsMultipleBOf(Double) As Boolean
   - Double.IsMultipleBOf(Double) As Boolean
   - Short.PercentageOf(Double) As Boolean
   - UShort.PercentageOf(Double) As Boolean
   - Integer.PercentageOf(Double) As Boolean
   - UInteger.PercentageOf(Double) As Boolean
   - Long.PercentageOf(Double) As Boolean
   - ULong.PercentageOf(Double) As Boolean
   - Decimal.PercentageOf(Double) As Boolean
   - Single.PercentageOf(Double) As Boolean
   - Double.PercentageOf(Double) As Boolean
   - Short.FractionBy(Double) As Double
   - UShort.FractionBy(Double) As Double
   - Integer.FractionBy(Double) As Double
   - UInteger.FractionBy(Double) As Double
   - Long.FractionBy(Double) As Double
   - ULong.FractionBy(Double) As Double
   - Decimal.FractionBy(Double) As Double
   - Single.FractionBy(Double) As Double
   - Double.FractionBy(Double) As Double
   - Short.Formatted(Double, Opt: Integer) As String
   - UShort.Formatted(Double, Opt: Integer) As String
   - Integer.Formatted(Double, Opt: Integer) As String
   - UInteger.Formatted(Double, Opt: Integer) As String
   - Long.Formatted(Double, Opt: Integer) As String
   - ULong.Formatted(Double, Opt: Integer) As String
   - Decimal.Formatted(Double, Opt: Integer) As String
   - Single.Formatted(Double, Opt: Integer) As String
   - Double.Formatted(Double, Opt: Integer) As String
   - Short.Formatted(Double, CultureInfo, Opt: Integer) As String
   - UShort.Formatted(Double, CultureInfo, Opt: Integer) As String
   - Integer.Formatted(Double, CultureInfo, Opt: Integer) As String
   - UInteger.Formatted(Double, CultureInfo, Opt: Integer) As String
   - Long.Formatted(Double, CultureInfo, Opt: Integer) As String
   - ULong.Formatted(Double, CultureInfo, Opt: Integer) As String
   - Decimal.Formatted(Double, CultureInfo, Opt: Integer) As String
   - Single.Formatted(Double, CultureInfo, Opt: Integer) As String
   - Double.Formatted(Double, CultureInfo, Opt: Integer) As String
   - Short.IsPrime() As Boolean
   - UShort.IsPrime() As Boolean
   - Integer.IsPrime() As Boolean
   - UInteger.IsPrime() As Boolean
   - Long.IsPrime() As Boolean
   - Short.ToVBHex() As String
   - UShort.ToVBHex() As String
   - Integer.ToVBHex() As String
   - UInteger.ToVBHex() As String
   - Long.ToVBHex() As String
   - Short.ToCSharpHex() As String
   - UShort.ToCSharpHex() As String
   - Integer.ToCSharpHex() As String
   - UInteger.ToCSharpHex() As String
   - Long.ToCSharpHex() As String
