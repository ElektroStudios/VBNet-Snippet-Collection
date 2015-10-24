' ***********************************************************************
' Author   : Elektro
' Modified : 24-October-2015
' ***********************************************************************
' <copyright file="Xml Extensions.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Functions "

' IEnumerable(Of XElement).DistinctByElement(string)
' IEnumerable(Of XElement).SortByElement(string)

' XDocument.DistinctByElement(string, string)
' XDocument.GetXPaths()
' XDocument.SortByElement(string, string)
' XDocument.ToXmlDocument()

' XmlDocument.DistinctByElement(string, string)
' XmlDocument.GetXPaths()
' XmlDocument.SortByElement(string, string)
' XmlDocument.ToXDocument()

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports Microsoft.VisualBasic
Imports System
Imports System.Diagnostics
Imports System.Runtime.CompilerServices
Imports System.Xml
Imports System.Xml.Linq

#End Region

#Region " Xml Extensions "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains custom extension methods to use with some of the <see cref="System.Xml"/> namespace members.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module XmlExtensions

#Region " Public Extension Methods "

#Region " Type Conversion "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Converts an <see cref="XmlDocument"/> to <see cref="XDocument"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xml As String =
    '''     <Songs>
    '''         <Song><Name>My Song 3.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>.ToString
    ''' 
    ''' Dim xmlDoc As New XmlDocument
    ''' xmlDoc.LoadXml(xml)
    ''' 
    ''' Dim xDoc As XDocument = xmlDoc.ToXDocument
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="XmlDocument"/>.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The <see cref="XDocument"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function ToXDocument(ByVal sender As XmlDocument) As XDocument

        Return XDocument.Parse(sender.InnerXml.TrimStart(ControlChars.Lf, " "c))

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Converts an <see cref="XDocument"/> to <see cref="XmlDocument"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xDoc As XDocument =
    '''     <?xml version="1.0" encoding="Windows-1252"?>
    '''     <!--XML Songs Database-->
    '''     <Songs>
    '''         <Song><Name>My Song 3.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>
    ''' 
    ''' Dim xmlDoc As XmlDocument = xDoc.ToXmlDocument
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="XDocument"/>.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The <see cref="XmlDocument"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function ToXmlDocument(ByVal sender As XDocument) As XmlDocument

        Dim xmlDoc As New XmlDocument
        xmlDoc.LoadXml(sender.ToString)
        Return xmlDoc

    End Function

#End Region

#Region " X-Path expressions "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a <see cref="IEnumerable(Of String)"/> collection with the avaliable XPath expressions of an <see cref="XmlDocument"/> document.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xml As String =
    '''     <Songs>
    '''         <Song><Name>My Song 3.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>.ToString
    ''' 
    ''' Dim xmlDoc As New XmlDocument
    ''' xmlDoc.LoadXml(xml)
    ''' 
    ''' Dim xPathList As IEnumerable(Of String) = xmlDoc.GetXPaths
    ''' MessageBox.Show(String.Join(Environment.NewLine, xPathList))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="XmlDocument"/>.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' A <see cref="IEnumerable(Of String)"/> collection with the avaliable XPath expressions.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function GetXPaths(ByVal sender As XmlDocument) As IEnumerable(Of String)

        ' Dim xmlReader As XmlReader = sender.CreateNavigator.ReadSubtree
        Dim nodeList As New List(Of String)
        Dim xPathList As New List(Of String)
        Dim xPath As String

        Using xmlReader As XmlReader = sender.CreateNavigator.ReadSubtree

            While xmlReader.Read

                If xmlReader.NodeType = XmlNodeType.Element Then

                    If nodeList.Count <= xmlReader.Depth Then
                        nodeList.Add(xmlReader.Name)
                    Else
                        nodeList(xmlReader.Depth) = xmlReader.Name
                    End If

                    xPath = String.Join("/", nodeList.ToArray(), 0, xmlReader.Depth + 1)

                    If Not xPathList.Contains(xPath) Then
                        xPathList.Add(xPath)
                    End If

                End If

            End While

            Return xPathList

        End Using

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a <see cref="IEnumerable(Of String)"/> collection with the avaliable XPath expressions of an <see cref="XDocument"/> document.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xDoc As XDocument =
    '''     <?xml version="1.0" encoding="Windows-1252"?>
    '''     <!--XML Songs Database-->
    '''     <Songs>
    '''         <Song><Name>My Song 3.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>
    ''' 
    ''' Dim xPathList As IEnumerable(Of String) = xDoc.GetXPaths
    ''' MessageBox.Show(String.Join(Environment.NewLine, xPathList))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="XDocument"/>.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' A <see cref="IEnumerable(Of String)"/> collection with the avaliable XPath expressions.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function GetXPaths(ByVal sender As XDocument) As IEnumerable(Of String)

        Return XmlExtensions.ToXmlDocument(sender).GetXPaths()

    End Function

#End Region

#Region " Duplicate Removal "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes duplicated values by the specified element of an <see cref="XDocument"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xDoc As XDocument =
    '''     <?xml version="1.0" encoding="Windows-1252"?>
    '''     <!--XML Songs Database-->
    '''     <Songs>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>
    ''' 
    ''' xDoc = xDoc.DistinctByElement(rootElementName:="Song", elementName:="Name")
    ''' MessageBox.Show(xDoc.ToString)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="XDocument"/>.
    ''' </param>
    ''' 
    ''' <param name="rootElementName">
    ''' The root Xml element name.
    ''' </param>
    ''' 
    ''' <param name="elementName">
    ''' The element name to remove its duplicated values.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The <see cref="XDocument"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function DistinctByElement(ByVal sender As XDocument,
                                      ByVal rootElementName As String,
                                      ByVal elementName As String) As XDocument

        sender.Root.Elements(rootElementName).
                    GroupBy(Function(xElement As XElement) xElement.Element(elementName).Value).
                    SelectMany(Function(group As IGrouping(Of String, XElement)) group.Skip(1)).
                    Remove()

        Return sender

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes duplicated values by the specified element of an <see cref="XmlDocument"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xml As String =
    '''     <Songs>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>.ToString
    ''' 
    ''' Dim xmlDoc As New XmlDocument
    ''' xmlDoc.LoadXml(xml)
    ''' 
    ''' xmlDoc = xmlDoc.DistinctByElement(rootElementName:="Song", elementName:="Name")
    ''' MessageBox.Show(xmlDoc.InnerXml)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="XmlDocument"/>.
    ''' </param>
    ''' 
    ''' <param name="rootElementName">
    ''' The root Xml element name.
    ''' </param>
    ''' 
    ''' <param name="elementName">
    ''' The element name to remove its duplicated values.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The <see cref="XmlDocument"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function DistinctByElement(ByVal sender As XmlDocument,
                                      ByVal rootElementName As String,
                                      ByVal elementName As String) As XmlDocument

        Return XmlExtensions.ToXmlDocument(sender.ToXDocument.DistinctByElement(rootElementName, elementName))

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes duplicated values by the specified element of an <see cref="IEnumerable(Of XElement)"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xDoc As XDocument =
    '''     <?xml version="1.0" encoding="Windows-1252"?>
    '''     <!--XML Songs Database-->
    '''     <Songs>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>
    ''' 
    ''' For Each el As XElement In xDoc.<Songs>.<Song>.DistinctByElement(elementName:="Name")
    '''     MessageBox.Show(el.Value)
    ''' Next
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="IEnumerable(Of XElement)"/>.
    ''' </param>
    ''' 
    ''' <param name="elementName">
    ''' The element name to remove its duplicated values.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The <see cref="IEnumerable(Of XElement)"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function DistinctByElement(ByVal sender As IEnumerable(Of XElement),
                                      ByVal elementName As String) As IEnumerable(Of XElement)

        sender.GroupBy(Function(xElement As XElement) xElement.Element(elementName).Value).
               SelectMany(Function(group As IGrouping(Of String, XElement)) group.Skip(1)).
               Remove()

        Return sender

    End Function

#End Region

#Region " Sorting "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Sorts an <see cref="XDocument"/> by the specified element.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xDoc As XDocument =
    '''     <?xml version="1.0" encoding="Windows-1252"?>
    '''     <!--XML Songs Database-->
    '''     <Songs>
    '''         <Song><Name>My Song 3.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>
    ''' 
    ''' xDoc = xDoc.SortByElement(rootElementName:="Song", elementName:="Name")
    ''' 
    ''' MessageBox.Show(xDoc.ToString)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="XDocument"/>.
    ''' </param>
    ''' 
    ''' <param name="rootElementName">
    ''' The root element name.
    ''' </param>
    ''' 
    ''' <param name="elementName">
    ''' The element name to sort by.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The sorted <see cref="XDocument"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function SortByElement(ByVal sender As XDocument,
                                  ByVal rootElementName As String,
                                  ByVal elementName As String) As XDocument

        sender.Root.ReplaceNodes(sender.Root.Elements(rootElementName).
                    OrderBy(Function(xElement As XElement) xElement.Element(elementName).Value))

        Return sender

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Sorts an <see cref="XmlDocument"/> by the specified element.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xml As String =
    '''     <Songs>
    '''         <Song><Name>My Song 3.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>.ToString
    ''' 
    ''' Dim xmlDoc As New XmlDocument
    ''' xmlDoc.LoadXml(xml)
    ''' 
    ''' xmlDoc = xmlDoc.SortByElement(rootElementName:="Song", elementName:="Name")
    ''' MessageBox.Show(xmlDoc.InnerXml)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="XmlDocument"/>.
    ''' </param>
    ''' 
    ''' <param name="rootElementName">
    ''' The root element name.
    ''' </param>
    ''' 
    ''' <param name="elementName">
    ''' The element name to sort by.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The sorted <see cref="XmlDocument"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function SortByElement(ByVal sender As XmlDocument,
                                  ByVal rootElementName As String,
                                  ByVal elementName As String) As XmlDocument

        Return XmlExtensions.ToXmlDocument(sender.ToXDocument.SortByElement(rootElementName, elementName))

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Sorts an <see cref="IEnumerable(Of XElement)"/> by the specified element.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim xml As String =
    '''     <Songs>
    '''         <Song><Name>My Song 3.mp3</Name></Song>
    '''         <Song><Name>My Song 1.mp3</Name></Song>
    '''         <Song><Name>My Song 2.mp3</Name></Song>
    '''     </Songs>.ToString
    ''' 
    ''' Dim xmlDoc As New XmlDocument
    ''' xmlDoc.LoadXml(xml)
    ''' 
    ''' xmlDoc = xmlDoc.SortByElement(rootElementName:="Song", elementName:="Name")
    ''' MessageBox.Show(xmlDoc.InnerXml)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="IEnumerable(Of XElement)"/>.
    ''' </param>
    ''' 
    ''' <param name="elementName">
    ''' The element name to sort by.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The sorted <see cref="IEnumerable(Of XElement)"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function SortByElement(ByVal sender As IEnumerable(Of XElement),
                                  ByVal elementName As String) As IEnumerable(Of XElement)

        Return sender.OrderBy(Function(xElement As XElement) xElement.Element(elementName).Value)

    End Function

#End Region

#End Region

End Module

#End Region
