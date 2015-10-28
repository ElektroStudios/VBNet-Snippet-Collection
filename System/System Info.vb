
' ***********************************************************************
' Author   : Elektro
' Modified : 28-October-2015
' ***********************************************************************
' <copyright file="System Info.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Child Classes "

' SystemInfo.OS
' SystemInfo.Programs

#End Region

#Region " Enumerations "

' SystemInfo.Architecture As Integer

#End Region

#Region " Properties "

' SystemInfo.OS.CurrentArchitecture As SystemInfo.Architecture
' SystemInfo.Programs.DefaultWebBrowser() As String
' SystemInfo.Programs.IExplorerVersion() As Version

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports Microsoft.Win32
Imports System
Imports System.Diagnostics
Imports System.Linq

#End Region

#Region " SystemInfo "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains related Windows operating system's information utilities.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module SystemInfo

#Region " Child Classes "

#Region " Operating System "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related operating system info.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class OS

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the architecture of the current operating system is 32 or 64 Bits.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' An <see cref="SystemInfo.Architecture"></see> object that specifies the architecture of the current operating system.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property CurrentArchitecture As Architecture
            <DebuggerStepThrough>
            Get
                Return If(Environment.Is64BitOperatingSystem,
                          Architecture.X64,
                          Architecture.X86)
            End Get
        End Property

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="SystemInfo.OS"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Hidden Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified System.Object instances are considered equal.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        Public Shadows Function Equals(ByVal obj As Object) As Boolean
            Return MyBase.Equals(obj)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified System.Object instances are the same instance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        Public Shadows Function ReferenceEquals(ByVal objA As Object, ByVal objB As Object) As Boolean
            Return Nothing
        End Function

#End Region

    End Class

#End Region

#Region " Programs "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related system's programs info.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class Programs

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the filepath of the default web-browser that is registered on the current operating system.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The filepath of the default web-browser that is registered on the current operating system.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property DefaultWebBrowser As String
            <DebuggerStepThrough>
            Get

                Dim regValue As String = String.Empty

                Using regKey As RegistryKey = Registry.ClassesRoot.OpenSubKey("HTTP\Shell\Open\Command", writable:=False)

                    regValue = regKey.GetValue(Nothing, String.Empty, RegistryValueOptions.None).ToString

                    regValue = regValue.Substring(0, regValue.LastIndexOf(".exe", StringComparison.OrdinalIgnoreCase) + ".exe".Length).
                                        Trim({ControlChars.Quote})

                End Using

                Return regValue
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the version of the Internet Explorer that is installed on the current operating system.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The version of the Internet Explorer that is installed on the current operating system.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property IExplorerVersion() As Version
            <DebuggerStepThrough>
            Get
                Return New Version(FileVersionInfo.GetVersionInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "ieframe.dll")).ProductVersion)
            End Get
        End Property

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="SystemInfo.Programs"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Hidden Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified System.Object instances are considered equal.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        Public Shadows Function Equals(ByVal obj As Object) As Boolean
            Return MyBase.Equals(obj)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified System.Object instances are the same instance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        Public Shadows Function ReferenceEquals(ByVal objA As Object, ByVal objB As Object) As Boolean
            Return Nothing
        End Function

#End Region

    End Class

#End Region

#End Region

#Region " Enumerations "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Specifies a Windows operating system architecture.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public Enum Architecture As Integer

        ''' <summary>
        ''' 32-Bit
        ''' </summary>
        X86 = 32

        ''' <summary>
        ''' 64-Bit
        ''' </summary>
        X64 = 64

    End Enum

#End Region

#Region " Hidden Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether the specified System.Object instances are considered equal.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Function Equals(ByVal obj As Object) As Boolean
        Return Nothing
    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether the specified System.Object instances are the same instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Function ReferenceEquals(ByVal objA As Object, ByVal objB As Object) As Boolean
        Return Nothing
    End Function

#End Region

End Module

#End Region
