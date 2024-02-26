' ***********************************************************************
' Author   : Elektro
' Modified : 26-October-2015
' ***********************************************************************
' <copyright file="Char Extensions.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Functions "

' Char.IsDiacritic As boolean

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
Imports System.Diagnostics
Imports System.Runtime.CompilerServices
Imports System.Text

#End Region

#Region " Char Extensions "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains custom extension methods to use with the <see cref="Char"/> type.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module CharExtensions

#Region " Public Extension Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether a character is diacritic or else contains a diacritical mark.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' MsgBox("á"c.IsDiacritic)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source character.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <c>true</c> if character is diacritic or contains a diacritical mark, <c>false</c> otherwise.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function IsDiacritic(ByVal sender As Char) As Boolean

        Dim descomposed As Char() = sender.ToString.Normalize(NormalizationForm.FormKD).ToCharArray
        Return (descomposed.Count <> 1 OrElse String.IsNullOrWhiteSpace(descomposed))

    End Function

#End Region

End Module

#End Region
