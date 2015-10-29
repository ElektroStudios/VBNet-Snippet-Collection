' ***********************************************************************
' Author   : Elektro
' Modified : 29-October-2015
' ***********************************************************************
' <copyright file="Logfile Writer.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Cosntructors "

' LogfileWriter.New(String, Opt: Encoding)

#End Region

#Region " Properties "

' LogfileWriter.Filepath As String
' LogfileWriter.Encoding As Encoding
' LogfileWriter.EntryFormat As String

#End Region

#Region " Methods "

' LogfileWriter.WriteEntry(TraceEventType, String)
' LogfileWriter.WriteText(String)
' LogfileWriter.WriteNewLine()
' LogfileWriter.Clear()

#End Region

#End Region

#Region " Usage Examples "

'Public Class Form1 : Inherits Form
'
'    Private logfile As New LogfileWriter(String.Format("{0}.log", My.Application.Info.AssemblyName)) With
'        {
'            .EntryFormat = "[{1}] | {2,-11} | {3}"
'        } ' {0}=Date, {1}=Time, {2}=Event, {3}=Message.
'
'    Private Sub Form1_Load() Handles MyBase.Load
'
'        With Me.logfile
'            .Clear()
'            .WriteText("#########################################")
'            .WriteNewLine()
'            .WriteText(String.Format("          Log Date {0}          ", DateTime.Now.Date.ToShortDateString))
'            .WriteNewLine()
'            .WriteText("#########################################")
'            .WriteNewLine()
'            .WriteEntry(TraceEventType.Information, "Application is being initialized.")
'        End With
'
'        Try
'            Dim setting As Integer = Integer.Parse(" Hello World! :D ")
'
'        Catch ex As Exception
'            Me.logfile.WriteEntry(TraceEventType.Critical, "Cannot parse 'setting' object in 'Sub Form1_Load()' method.")
'            Me.logfile.WriteEntry(TraceEventType.Information, "Exiting...")
'            Application.Exit()
'
'        End Try
'
'    End Sub
'
'End Class

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Text

#End Region

#Region " Logfile Writer "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' A simple logging system assistant that helps to create a log for the current application.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
''' <example> This is a code example.
''' <code>
''' Public Class Form1 : Inherits Form
''' 
'''     Private logfile As New LogfileWriter(String.Format("{0}.log", My.Application.Info.AssemblyName)) With
'''         {
'''             .EntryFormat = "[{1}] | {2,-11} | {3}"
'''         } ' {0}=Date, {1}=Time, {2}=Event, {3}=Message.
''' 
'''     Private Sub Form1_Load() Handles MyBase.Load
''' 
'''         With Me.logfile
'''             .Clear()
'''             .WriteText("#########################################")
'''             .WriteNewLine()
'''             .WriteText(String.Format("          Log Date {0}          ", DateTime.Now.Date.ToShortDateString))
'''             .WriteNewLine()
'''             .WriteText("#########################################")
'''             .WriteNewLine()
'''             .WriteEntry(TraceEventType.Information, "Application is being initialized.")
'''         End With
''' 
'''         Try
'''             Dim setting As Integer = Integer.Parse(" Hello World! :D ")
''' 
'''         Catch ex As Exception
'''             Me.logfile.WriteEntry(TraceEventType.Critical, "Cannot parse 'setting' object in 'Sub Form1_Load()' method.")
'''             Me.logfile.WriteEntry(TraceEventType.Information, "Exiting...")
'''             Application.Exit()
''' 
'''         End Try
''' 
'''     End Sub
''' 
''' End Class
''' </code>
''' </example>
''' ----------------------------------------------------------------------------------------------------
Public NotInheritable Class LogfileWriter : Implements IDisposable

#Region " Properties "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the log filepath.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <value>
    ''' The log filepath.
    ''' </value>
    ''' ----------------------------------------------------------------------------------------------------
    Public ReadOnly Property Filepath As String
        Get
            Return Me.filepathB
        End Get
    End Property
    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' ( Backing field )
    ''' The log filepath.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private ReadOnly filepathB As String

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the logfile encoding.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <value>
    ''' The logfile encoding.
    ''' </value>
    ''' ----------------------------------------------------------------------------------------------------
    Public ReadOnly Property Encoding As Encoding
        Get
            Return Me.encodingB
        End Get
    End Property
    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' ( Backing field )
    ''' The logfile encoding.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private ReadOnly encodingB As Encoding

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the format of a log entry.
    ''' {0}=Date, {1}=Time, {2}=Event, {3}=Message.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <value>
    ''' The format of a log entry.
    ''' </value>
    ''' ----------------------------------------------------------------------------------------------------
    Public Property EntryFormat As String = "[{0}] [{1}] | {2,-11} | {3}"

#End Region

#Region " Variables "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' The <see cref="StreamWriter"/> where is written the logging data
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private sw As StreamWriter

#End Region

#Region " Constructors "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Prevents a default instance of the <see cref="LogfileWriter"/> class from being created.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub New()
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Initializes a new instance of the <see cref="LogfileWriter"/> class.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="filepath">
    ''' The log filepath.
    ''' </param>
    ''' 
    ''' <param name="Encoding">
    ''' The file encoding.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub New(ByVal filepath As String, Optional ByVal encoding As Encoding = Nothing)

        If encoding Is Nothing Then
            encoding = System.Text.Encoding.Default
        End If

        Me.filepathB = filepath
        Me.encodingB = encoding
        Me.sw = New StreamWriter(filepath, append:=True, encoding:=encoding, bufferSize:=4096) With {.AutoFlush = True}

    End Sub

#End Region

#Region " Public Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Writes a new entry on the logfile.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="eventType">
    ''' The type of event.
    ''' </param>
    ''' 
    ''' <param name="message">
    ''' The message to log.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub WriteEntry(ByVal eventType As TraceEventType,
                          ByVal message As String)

        Dim localDate As String = DateTime.Now.Date.ToShortDateString
        Dim localTime As String = DateTime.Now.ToLongTimeString

        Me.sw.WriteLine(String.Format(Me.EntryFormat, localDate, localTime, eventType.ToString, message))

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Writes any text on the logfile.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="text">
    ''' The text to write.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub WriteText(ByVal text As String)

        Me.sw.WriteLine(text)

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Writes an empty line on the logfile.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub WriteNewLine()

        Me.sw.WriteLine()

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Clears the logfile content.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub Clear()

        Me.sw.Close()

        Me.sw = New StreamWriter(Me.filepathB, append:=False, encoding:=Me.encodingB, bufferSize:=4096) With {.AutoFlush = True}
        Me.sw.BaseStream.SetLength(0)

    End Sub

#End Region

#Region " IDisposable Implementation "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' To detect redundant calls when disposing.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private isDisposed As Boolean = False

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Releases all the resources used by this instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub Dispose() Implements IDisposable.Dispose
        Me.Dispose(isDisposing:=True)
        GC.SuppressFinalize(obj:=Me)
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    ''' Releases unmanaged and - optionally - managed resources.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="isDisposing">
    ''' <see langword="True"/>  to release both managed and unmanaged resources; 
    ''' <see langword="False"/> to release only unmanaged resources.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Protected Sub Dispose(ByVal isDisposing As Boolean)

        If (Not Me.isDisposed) AndAlso (isDisposing) Then

            If (Me.sw IsNot Nothing) Then
                Me.sw.Close()
            End If

        End If

        Me.isDisposed = True

    End Sub

#End Region

End Class

#End Region
