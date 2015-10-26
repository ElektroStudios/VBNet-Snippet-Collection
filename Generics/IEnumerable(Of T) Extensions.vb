' ***********************************************************************
' Author   : Elektro
' Modified : 26-October-2015
' ***********************************************************************
' <copyright file="IEnumerableExtensions.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Functions "

' IEnumerable(Of T)().ConcatMultiple(IEnumerable(Of T)()) As IEnumerable(Of T)
' IEnumerable(Of T)().StringJoin As IEnumerable(Of T)
' IEnumerable(Of T).CountEmptyItems As Integer
' IEnumerable(Of T).CountNonEmptyItems As Integer
' IEnumerable(Of T).Duplicates As IEnumerable(Of T)
' IEnumerable(Of T).Randomize As IEnumerable(Of T)
' IEnumerable(Of T).RemoveDuplicates As IEnumerable(Of T)
' IEnumerable(Of T).SplitIntoNumberOfElements(Integer) As IEnumerable(Of T)
' IEnumerable(Of T).SplitIntoNumberOfElements(Integer, Boolean, T) As IEnumerable(Of T)
' IEnumerable(Of T).SplitIntoParts(Integer) As IEnumerable(Of T)
' IEnumerable(Of T).UniqueDuplicates As IEnumerable(Of T)
' IEnumerable(Of T).Uniques As IEnumerable(Of T)

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

#Region " IEnumerable(Of T) Extensions "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains custom extension methods to use with an <see cref="IEnumerable(Of T)"/>.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module IEnumerableExtensions

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets all the duplicated values of the source <see cref="IEnumerable(Of T)"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim col As IEnumerable(Of Integer) = {1, 1, 2, 2, 3, 3, 0}
    ''' Debug.WriteLine(String.Join(", ", col.Duplicates))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collection.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of T)"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function Duplicates(Of T)(ByVal sender As IEnumerable(Of T)) As IEnumerable(Of T)

        Return sender.GroupBy(Function(value As T) value).
                      Where(Function(group As IGrouping(Of T, T)) group.Count > 1).
                      SelectMany(Function(group As IGrouping(Of T, T)) group)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the unique duplicated values of the source <see cref="IEnumerable(Of T)"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim col As IEnumerable(Of Integer) = {1, 1, 2, 2, 3, 3, 0}
    ''' Debug.WriteLine(String.Join(", ", col.UniqueDuplicates))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collection.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of T)"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function UniqueDuplicates(Of T)(ByVal sender As IEnumerable(Of T)) As IEnumerable(Of T)

        Return sender.GroupBy(Function(value As T) value).
                      Where(Function(group As IGrouping(Of T, T)) group.Count > 1).
                      Select(Function(group As IGrouping(Of T, T)) group.Key)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the unique values of the source <see cref="IEnumerable(Of T)"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim col As IEnumerable(Of Integer) = {1, 1, 2, 2, 3, 3, 0}
    ''' Debug.WriteLine(String.Join(", ", col.Uniques))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collection.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of T)"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function Uniques(Of T)(ByVal sender As IEnumerable(Of T)) As IEnumerable(Of T)

        Return sender.Except(IEnumerableExtensions.UniqueDuplicates(sender))

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Removes duplicated values in the source <see cref="IEnumerable(Of T)"/>.
    ''' </summary>   
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim col As IEnumerable(Of Integer) = {1, 1, 2, 2, 3, 3, 0}
    ''' Debug.WriteLine(String.Join(", ", col.RemoveDuplicates))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collection.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of T)"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function RemoveDuplicates(Of T)(ByVal sender As IEnumerable(Of T)) As IEnumerable(Of T)

        Return sender.Distinct

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Splits the source <see cref="IEnumerable(Of T)"/> into the specified amount of secuences.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    '''  Dim mainCol As IEnumerable(Of Integer) = {1, 2, 3, 4, 5, 6, 7, 8, 9, 0}
    '''  Dim splittedCols As IEnumerable(Of IEnumerable(Of Integer)) = mainCol.SplitIntoParts(amount:=2)
    '''  splittedCols.ToList.ForEach(Sub(col As IEnumerable(Of Integer))
    '''                                  Debug.WriteLine(String.Join(", ", col))
    '''                              End Sub)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collection.
    ''' </param>
    ''' 
    ''' <param name="amount">
    ''' The target amount of secuences.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of IEnumerable(Of T))"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function SplitIntoParts(Of T)(ByVal sender As IEnumerable(Of T),
                                         ByVal amount As Integer) As IEnumerable(Of IEnumerable(Of T))

        If (amount = 0) OrElse (amount > sender.Count) OrElse (sender.Count Mod amount <> 0) Then
            Throw New ArgumentOutOfRangeException(paramName:="amount",
                                                  message:="value should be greater than '0', smallest than 'col.Count', and multiplier of 'col.Count'.")
        End If

        Dim chunkSize As Integer = CInt(Math.Ceiling(sender.Count() / amount))

        Return From index As Integer In Enumerable.Range(0, amount)
               Select sender.Skip(chunkSize * index).Take(chunkSize)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Splits the source <see cref="IEnumerable(Of T)"/> into secuences with the specified amount of elements.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    '''  Dim mainCol As IEnumerable(Of Integer) = {1, 2, 3, 4, 5, 6, 7, 8, 9}
    '''  Dim splittedCols As IEnumerable(Of IEnumerable(Of Integer)) = mainCol.SplitIntoNumberOfElements(amount:=4)
    '''  splittedCols.ToList.ForEach(Sub(col As IEnumerable(Of Integer))
    '''                                  Debug.WriteLine(String.Join(", ", col))
    '''                              End Sub)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collection.
    ''' </param>
    ''' 
    ''' <param name="amount">
    ''' The target amount of elements.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of IEnumerable(Of T))"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function SplitIntoNumberOfElements(Of T)(ByVal sender As IEnumerable(Of T),
                                                    ByVal amount As Integer) As IEnumerable(Of IEnumerable(Of T))

        Return From index As Integer In Enumerable.Range(0, CInt(Math.Ceiling(sender.Count() / amount)))
               Select sender.Skip(index * amount).Take(amount)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Splits the source <see cref="IEnumerable(Of T)"/> into secuences with the specified amount of elements.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    '''  Dim mainCol As IEnumerable(Of Integer) = {1, 2, 3, 4, 5, 6, 7, 8, 9}
    '''  Dim splittedCols As IEnumerable(Of IEnumerable(Of Integer)) = mainCol.SplitIntoNumberOfElements(amount:=4, fillEmpty:=True, valueToFill:=0)
    '''  splittedCols.ToList.ForEach(Sub(col As IEnumerable(Of Integer))
    '''                                  Debug.WriteLine(String.Join(", ", col))
    '''                              End Sub)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collection.
    ''' </param>
    ''' 
    ''' <param name="amount">
    ''' The target amount of elements.
    ''' </param>
    ''' 
    ''' <param name="fillEmpty">
    ''' If set to <c>true</c>, generates empty elements to fill the last secuence's part amount.
    ''' </param>
    ''' 
    ''' <param name="valueToFill">
    ''' An optional value used to fill the last secuence's part amount.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of IEnumerable(Of T))"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function SplitIntoNumberOfElements(Of T)(ByVal sender As IEnumerable(Of T),
                                                    ByVal amount As Integer,
                                                    ByVal fillEmpty As Boolean,
                                                    Optional valueToFill As T = Nothing) As IEnumerable(Of IEnumerable(Of T))

        Return (From count As Integer In Enumerable.Range(0, CInt(Math.Ceiling(sender.Count() / amount)))).
                Select(Function(count)

                           Select Case fillEmpty

                               Case True
                                   If (sender.Count - (count * amount)) >= amount Then
                                       Return sender.Skip(count * amount).Take(amount)

                                   Else
                                       Return sender.Skip(count * amount).Take(amount).
                                                  Concat(Enumerable.Repeat(Of T)(
                                                         valueToFill,
                                                         amount - (sender.Count() - (count * amount))))
                                   End If

                               Case Else
                                   Return sender.Skip(count * amount).Take(amount)

                           End Select

                       End Function)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Randomizes the elements of the source <see cref="IEnumerable(Of T)"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim col As IEnumerable(Of Integer) = {1, 2, 3, 4, 5, 6, 7, 8, 9}
    ''' Debug.WriteLine(String.Join(", ", col.Randomize))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collection.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of T)"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function Randomize(Of T)(ByVal sender As IEnumerable(Of T)) As IEnumerable(Of T)

        Dim rand As New Random

        Return From item As T In sender
               Order By rand.Next

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Concatenates multiple <see cref="IEnumerable(Of T)"/> at once into a single <see cref="IEnumerable(Of T)"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim col1 As IEnumerable(Of Integer) = {1, 2, 3}
    ''' Dim col2 As IEnumerable(Of Integer) = {4, 5, 6}
    ''' Dim col3 As IEnumerable(Of Integer) = {7, 8, 9}
    ''' Debug.WriteLine(String.Join(", ", {col1, col2, col3}.ConcatMultiple))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    ''' 
    ''' <param name="sender">
    ''' The source collections.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="IEnumerable(Of T)"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function ConcatMultiple(Of T)(ByVal sender As IEnumerable(Of T)()) As IEnumerable(Of T)

        Return sender.SelectMany(Function(col As IEnumerable(Of T)) col)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Joins multiple <see cref="IEnumerable(Of T)"/> at once into a single string.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim col1 As IEnumerable(Of Integer) = {1, 2, 3}
    ''' Dim col2 As IEnumerable(Of Integer) = {4, 5, 6}
    ''' Dim col3 As IEnumerable(Of Integer) = {7, 8, 9}
    ''' Debug.WriteLine({col1, col2, col3}.StringJoin(", ")))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <typeparam name="T">
    ''' </typeparam>
    '''     
    ''' <param name="separator">
    ''' The string to use as a separator.
    ''' </param>
    ''' 
    ''' <param name="sender">
    ''' The source collections.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see cref="String"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function StringJoin(Of T)(ByVal sender As IEnumerable(Of T)(),
                                     ByVal separator As String) As String

        Dim sb As New System.Text.StringBuilder

        For Each col As IEnumerable(Of T) In sender
            sb.Append(String.Join(separator, col) & separator)
        Next col

        Return sb.Remove(sb.Length - separator.Length, separator.Length).ToString

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim emptyItemCount As Integer = {"Hello", "   ", "World!"}.CountEmptyItems
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Counts the empty items of the source <see cref="IEnumerable(Of T)"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="IEnumerable(Of T)"/>.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The total amount of empty items.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function CountEmptyItems(Of T)(ByVal sender As IEnumerable(Of T)) As Integer

        Return (From item As T In sender
                Where (item.Equals(Nothing))).Count

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Counts the non-empty items of the source <see cref="IEnumerable(Of T)"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim nonEmptyItemCount As Integer = {"Hello", "   ", "World!"}.CountNonEmptyItems
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="sender">
    ''' The source <see cref="IEnumerable(Of T)"/>.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The total amount of non-empty items.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    <Extension>
    Public Function CountNonEmptyItems(Of T)(ByVal sender As IEnumerable(Of T)) As Integer

        Return (sender.Count - IEnumerableExtensions.CountEmptyItems(sender))

    End Function

End Module

#End Region
