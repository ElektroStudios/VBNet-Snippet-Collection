' ***********************************************************************
' Author   : Elektro
' Modified : 28-October-2015
' ***********************************************************************
' <copyright file="Array Extensions.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Functions "

' T().Resize As T()

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

#End Region

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains custom extension methods to use with an <see cref="Array"/>.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module ArrayExtensions

#Region " Public Extension Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Resizes the number of elements of the source <see cref="Array"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim myArray(50) As Integer
    ''' Console.WriteLine(String.Format("{0,-12}: {1}", "Initial Size", myArray.Length))
    ''' 
    ''' myArray = myArray.Resize(myArray.Length - 51)
    ''' Console.WriteLine(String.Format("{0,-12}: {1}", "New Size", myArray.Length))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' The array <see cref="Type"/>.
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source <see cref="Array"/>.
    ''' </param>
    ''' 
    ''' <param name="newSize">
    ''' The new size.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The resized <see cref="Array"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <exception cref="System.ArgumentOutOfRangeException">
    ''' newSize;Value greater than 0 is required.
    ''' </exception>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function Resize(Of T)(ByVal sender As T(),
                                 ByVal newSize As Integer) As T()

        If (newSize <= 0) Then
            Throw New System.ArgumentOutOfRangeException(paramName:="newSize", message:="Value greater than 0 is required.")
        End If

        Dim preserveLength As Integer = Math.Min(sender.Length, newSize)

        If (preserveLength > 0) Then
            Dim newArray As Array = Array.CreateInstance(sender.GetType.GetElementType, newSize)
            Array.Copy(sender, newArray, preserveLength)
            Return DirectCast(newArray, T())

        Else
            Return sender

        End If

    End Function

#End Region

End Module
