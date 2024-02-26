' ***********************************************************************
' Author   : Elektro
' Modified : 28-October-2015
' ***********************************************************************
' <copyright file="IEnumerable(Of Byte) Extensions.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Functions "

' IEnumerable(Of Byte).ToString(Encoding) As String
' Byte().ToString(Encoding) As String

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
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.CompilerServices

#End Region

#Region " IEnumerable(Of Byte) Extensions "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains custom extension methods to use with an <see cref="IEnumerable(Of Byte)"/>.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module IEnumerableOfByteExtensions

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a byte sequence to its String representation using the specified character <see cref="Encoding"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' MessageBox.Show(New Byte() {84, 101, 115, 116}.ToString(Encoding.Default))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="Array"/>.
    ''' </param>
    ''' 
    ''' <param name="encoding">
    ''' The character <see cref="Encoding"/> to decode the bytes.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The String representation.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function ToString(ByVal sender As IEnumerable(Of Byte),
                             ByVal encoding As Encoding) As String

        Return IEnumerableOfByteExtensions.ToString(sender.ToArray, encoding)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a byte sequence to its String representation using the specified character <see cref="Encoding"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' MessageBox.Show(New Byte() {84, 101, 115, 116}.ToString(Encoding.Default))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="Array"/>.
    ''' </param>
    ''' 
    ''' <param name="encoding">
    ''' The character <see cref="Encoding"/> to decode the bytes.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The String representation.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function ToString(ByVal sender As Byte(),
                             ByVal encoding As Encoding) As String

        If encoding Is Nothing Then
            encoding = System.Text.Encoding.Default
        End If

        Return encoding.GetString(sender)

    End Function

End Module

#End Region
