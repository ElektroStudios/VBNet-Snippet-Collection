' ***********************************************************************
' Author   : Elektro
' Modified : 26-October-2015
' ***********************************************************************
' <copyright file="RegEx Util.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Constants "

' RegExUtil.Patterns.AlphabeticText As String
' RegExUtil.Patterns.AlphanumericText As String
' RegExUtil.Patterns.CreditCard As String
' RegExUtil.Patterns.EMail As String
' RegExUtil.Patterns.Hex As String
' RegExUtil.Patterns.HtmlTag As String
' RegExUtil.Patterns.Ipv4 As String
' RegExUtil.Patterns.Ipv6 As String
' RegExUtil.Patterns.NumericText As String
' RegExUtil.Patterns.Phone As String
' RegExUtil.Patterns.SafeText As String
' RegExUtil.Patterns.Url As String
' RegExUtil.Patterns.USphone As String
' RegExUtil.Patterns.USssn As String
' RegExUtil.Patterns.USstate As String
' RegExUtil.Patterns.USzip As String

#End Region

#Region " Child Classes "

' RegExUtil.Patterns

#End Region

#Region " Types "

' RegExUtil.MatchPositionInfo <Serializable>

#End Region

#Region " Constructors "

' RegExUtil.MatchPositionInfo.New(String, Integer)

#End Region

#Region " Properties "

' RegExUtil.MatchPositionInfo.Text As String
' RegExUtil.MatchPositionInfo.StartIndex As Integer
' RegExUtil.MatchPositionInfo.EndIndex As Integer
' RegExUtil.MatchPositionInfo.Length As Integer

#End Region

#Region " Functions "

' RegExUtil.GetMatchesPositions(Regex, String, Integer) As IEnumerable(Of RegExUtil.MatchPositionInfo) 
' RegExUtil.Validate(String, Boolean) As Boolean

#End Region

#End Region

#Region " Usage Examples "

#Region " RegExUtil.MatchPositionInfo "

'Sub Test()
'
'    Dim regExpr As New Regex("Dog(s)?", RegexOptions.IgnoreCase)
'
'    Dim text As String = "One Dog!, Two Dogs!, three Dogs!"
'    RichTextBox1.Text = text
'
'    Dim matchesPos As IEnumerable(Of RegExUtil.MatchPositionInfo) = RegExUtil.GetMatchesPositions(regExpr, text, groupIndex:=0)
'
'    For Each matchPos As RegExUtil.MatchPositionInfo In matchesPos
'
'        Console.WriteLine(text.Substring(matchPos.StartIndex, matchPos.Length))
'
'        With RichTextBox1
'            .SelectionStart = matchPos.StartIndex
'            .SelectionLength = matchPos.Length
'            .SelectionBackColor = Color.IndianRed
'            .SelectionColor = Color.WhiteSmoke
'            .SelectionFont = New Font(RichTextBox1.Font.Name, RichTextBox1.Font.SizeInPoints, FontStyle.Bold)
'        End With
'
'    Next matchPos
'
'    With RichTextBox1
'        .SelectionStart = 0
'        .SelectionLength = 0
'    End With
'
'End Sub

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions

#End Region

#Region " RegEx Util "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains related RegEx utilities.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public NotInheritable Class RegExUtil

#Region " Types "

#Region " MatchPositionInfo "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Encapsulates a text value captured by a RegEx, with its start/end index.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <Serializable>
    Public Structure MatchPositionInfo

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the text value.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The text value.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property Text As String
            Get
                Return Me.textB
            End Get
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing field )
        ''' The text value.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private ReadOnly textB As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the start index.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The start index.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property StartIndex As Integer
            Get
                Return Me.startIndexB
            End Get
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing field )
        ''' The start index.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private ReadOnly startIndexB As Integer

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the end index.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The end index.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property EndIndex As Integer
            Get
                Return Me.endIndexB
            End Get
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing field )
        ''' The end index.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private ReadOnly endIndexB As Integer

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the matched text length.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The matched text length.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property Length As Integer
            Get
                Return Me.endIndexB - Me.startIndexB
            End Get
        End Property

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="MatchPositionInfo"/> structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="text">
        ''' The text value.
        ''' </param>
        ''' 
        ''' <param name="startIndex">
        ''' The start index.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New(ByVal text As String,
                       ByVal startIndex As Integer)

            Me.textB = text
            Me.startIndexB = startIndex
            Me.endIndexB = (startIndex + text.Length)

        End Sub

#End Region

    End Structure

#End Region

#End Region

#Region " Child Classes "

#Region " Patterns "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' A class that exposes common RegEx patterns.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public Class Patterns

#Region " Constants "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches an URL.
        ''' 
        ''' For Example:
        ''' http://url
        ''' ftp://url
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const Url As String =
            "^((((https?|ftps?|gopher|telnet|nntp)://)|(mailto:|news:))(%[0-9A-Fa-f]{2}|[-()_.!~*';/?:@&=+$,A-Za-z0-9])+)([).!';/?:,][[:blank:]])?$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches the content of an Html enclosed tag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const HtmlTag As String =
            "/^<([a-z]+)([^<]+)*(?:>(.*)<\/\1>|\s+\/>)$/"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches an IPv4 address.
        ''' 
        ''' For Example:
        ''' 127.0.0.1
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const Ipv4 As String =
            "/^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$/"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches an IPv6 address.
        ''' 
        ''' For Example:
        ''' FE80:0000:0000:0000:0202:B3FF:FE1E:8329
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const Ipv6 As String =
            "(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]).){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]).){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches a valid e-mail address.
        ''' 
        ''' For Example:
        ''' 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const EMail As String =
            "^[a-zA-Z0-9+&*-]+(?:\.[a-zA-Z0-9_+&*-]+)*@(?:[a-zA-Z0-9-]+\.)+[a-zA-Z]{2,7}$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches lower and upper case letters and all digits.
        ''' 
        ''' For Example:
        ''' 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const SafeText As String =
            "^[a-zA-Z0-9 .-]+$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches a valid credit card number, VISA or also a passport.
        ''' 
        ''' For Example:
        ''' 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const CreditCard As String =
            "^((4\d{3})|(5[1-5]\d{2})|(6011)|(7\d{3}))-?\d{4}-?\d{4}-?\d{4}|3[4,7]\d{13}$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches an United States zip code with optional dash-four.
        ''' 
        ''' For Example:
        ''' 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const USzip As String =
            "^\d{5}(-\d{4})?$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches an United States phone number with or without dashes.
        ''' 
        ''' For Example:
        ''' 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const USphone As String =
            "^\D?(\d{3})\D?\D?(\d{3})\D?(\d{4})$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches a 2 letter United States state abbreviations.
        ''' 
        ''' For Example:
        ''' 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const USstate As String =
            "^(AE|AL|AK|AP|AS|AZ|AR|CA|CO|CT|DE|DC|FM|FL|GA|GU|HI|ID|IL|IN|IA|KS|KY|LA|ME|MH|MD|MA|MI|MN|MS|MO|MP|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PW|PA|PR|RI|SC|SD|TN|TX|UT|VT|VI|VA|WA|WV|WI|WY)$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches a 9 digit United States social security number with dashes.
        ''' 
        ''' For Example:
        ''' 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const USssn As String =
            "[0-9]\{3\}-[0-9]\{2\}-[0-9]\{4\}"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches Hexadecimal values.
        ''' 
        ''' For Example:
        ''' #a3c113
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const Hex As String =
            "/^#?([a-f0-9]{6}|[a-f0-9]{3})$/"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches a Phone number.
        ''' 
        ''' Number in the following form: (###) ###-####
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const Phone As String =
            "/^(?[0-9]{3})?|[0-9]{3}[-. ]? [0-9]{3}[-. ]?[0-9]{4}$/"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches alphabetic text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const AlphabeticText As String =
            "^[A-Za-z]+$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches Alphanumeric text.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const AlphanumericText As String =
            "^[A-Za-z0-9]+$"

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A pattern that matches numeric text, integer or decimal.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Const NumericText As String =
            "^((4\d{3})|(5[1-5]\d{2})|(6011))-?\d{4}-?\d"

#End Region

    End Class

#End Region

#End Region

#Region " Public Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Validates the specified regular expression pattern.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="pattern">
    ''' The RegEx pattern.
    ''' </param>
    ''' 
    ''' <param name="ignoreErrors">
    ''' If set to<see langword="True"/>, ignore validation errors, otherwise, throws an exception if validation fails.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see langword="True"/> if pattern validation success, <see langword="False"/> otherwise.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Shared Function Validate(ByVal pattern As String,
                                    Optional ByVal ignoreErrors As Boolean = True) As Boolean

        Try
            Dim regEx As New Regex(pattern:=pattern)
            Return True

        Catch ex As Exception
            If Not ignoreErrors Then
                Throw
            End If
            Return False

        End Try

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Validates the specified regular expression pattern.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Sub Test()
    ''' 
    '''     Dim regExpr As New Regex("Dog(s)?", RegexOptions.IgnoreCase)
    ''' 
    '''     Dim text As String = "One Dog!, Two Dogs!, three Dogs!"
    '''     RichTextBox1.Text = text
    ''' 
    '''     Dim matchesPos As IEnumerable(Of RegExUtil.MatchPositionInfo) = RegExUtil.GetMatchesPositions(regExpr, text, groupIndex:=0)
    ''' 
    '''     For Each matchPos As RegExUtil.MatchPositionInfo In matchesPos
    ''' 
    '''         Console.WriteLine(text.Substring(matchPos.StartIndex, matchPos.Length))
    ''' 
    '''         With RichTextBox1
    '''             .SelectionStart = matchPos.StartIndex
    '''             .SelectionLength = matchPos.Length
    '''             .SelectionBackColor = Color.IndianRed
    '''             .SelectionColor = Color.WhiteSmoke
    '''             .SelectionFont = New Font(RichTextBox1.Font.Name, RichTextBox1.Font.SizeInPoints, FontStyle.Bold)
    '''         End With
    ''' 
    '''     Next matchPos
    ''' 
    '''     With RichTextBox1
    '''         .SelectionStart = 0
    '''         .SelectionLength = 0
    '''     End With
    ''' 
    ''' End Sub
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="regEx">
    ''' The RegEx pattern.
    ''' </param>
    ''' 
    ''' <param name="text">
    ''' If set to <see langword="True"/>, ignore validation errors, otherwise, throws an exception if validation fails.
    ''' </param>
    ''' 
    ''' <param name="groupIndex">
    ''' If set to <see langword="True"/>, ignore validation errors, otherwise, throws an exception if validation fails.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see langword="True"/> if pattern validation success, <see langword="False"/> otherwise.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Shared Iterator Function GetMatchesPositions(ByVal regEx As Regex,
                                                        ByVal text As String,
                                                        Optional ByVal groupIndex As Integer = 0) As IEnumerable(Of MatchPositionInfo)

        Dim match As Match = regEx.Match(text)

        Do While match.Success

            Yield New MatchPositionInfo(text:=match.Groups(groupIndex).Value,
                                        startIndex:=match.Groups(groupIndex).Index)

            match = match.NextMatch

        Loop

    End Function

#End Region

End Class

#End Region
