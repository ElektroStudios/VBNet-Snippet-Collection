' ***********************************************************************
' Author   : Elektro
' Modified : 03-November-2015
' ***********************************************************************
' <copyright file="IPC Util.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Required assembly references "

' Microsoft Windows UI Automation: 
' UIAutomationClient (UIAutomationClient.dll)
' UIAutomationTypes  (UIAutomationTypes.dll)

#End Region

#Region " Public Members Summary "

#Region " Child Classes "

' IpcUtil.SharedMemory
' IpcUtil.UIAutomation

#End Region

#Region " Functions "

' IpcUtil.SharedMemory.Create(String, Integer, Opt: MemoryMappedFileAccess) As MemoryMappedFile
' IpcUtil.SharedMemory.Read(MemoryMappedFile, Long, Long) As Byte()
' IpcUtil.SharedMemory.Read(String, Long, Long) As Byte()
' IpcUtil.SharedMemory.ReadAt(MemoryMappedFile, Long) As Byte
' IpcUtil.SharedMemory.ReadAt(String, Long) As Byte
' IpcUtil.SharedMemory.ReadCharAt(MemoryMappedFile, Long) As Char
' IpcUtil.SharedMemory.ReadCharAt(String, Long) As Char
' IpcUtil.SharedMemory.ReadString(MemoryMappedFile, Long, Long, Opt: Encoding) As String
' IpcUtil.SharedMemory.ReadString(String, Long, Long, Opt: Encoding) As String
' IpcUtil.SharedMemory.ReadStringToEnd(MemoryMappedFile, Opt: Encoding) As String
' IpcUtil.SharedMemory.ReadStringToEnd(String, Opt: Encoding) As String
' IpcUtil.SharedMemory.ReadToEnd(MemoryMappedFile) As Byte()
' IpcUtil.SharedMemory.ReadToEnd(String) As Byte()

' IpcUtil.UIAutomation.GetTitlebarText(IntPtr) As String
' IpcUtil.UIAutomation.SendText(IntPtr) As String
' IpcUtil.SetText(IntPtr, String) As Boolean
' IpcUtil.AppendText(IntPtr, String) As Boolean
' IpcUtil.InsertText(IntPtr, Integer, String) As Boolean

#End Region

#Region " Methods "

' IpcUtil.SharedMemory.Clear(MemoryMappedFile)
' IpcUtil.SharedMemory.Clear(String)
' IpcUtil.SharedMemory.Write(MemoryMappedFile, Byte())
' IpcUtil.SharedMemory.Write(String, Byte())
' IpcUtil.SharedMemory.WriteAt(MemoryMappedFile, Byte(), Long)
' IpcUtil.SharedMemory.WriteAt(MemoryMappedFile, Byte, Long)
' IpcUtil.SharedMemory.WriteAt(String, Byte(), Long)
' IpcUtil.SharedMemory.WriteAt(String, Byte, Long)
' IpcUtil.SharedMemory.WriteCharAt(MemoryMappedFile, Char, Long, Opt: Encoding)
' IpcUtil.SharedMemory.WriteCharAt(String, Char, Long, Opt: Encoding)
' IpcUtil.SharedMemory.WriteString(MemoryMappedFile, String, Opt: Encoding)
' IpcUtil.SharedMemory.WriteString(String, String, Opt: Encoding)
' IpcUtil.SharedMemory.WriteStringAt(MemoryMappedFile, String, Long, Opt: Encoding)
' IpcUtil.SharedMemory.WriteStringAt(String, String, Long, Opt: Encoding)

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
Imports System.IO.MemoryMappedFiles
Imports System.Text
Imports System.Windows.Automation
' Imports System.ComponentModel
' Imports System.Linq.Expressions
' Imports System.Reflection
' Imports System.Runtime.InteropServices

#End Region

#Region " IpcUtil "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains related Inter-process communication (IPC) and UI automation utilities.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module IpcUtil

#Region " P/Invoking "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Platform Invocation methods (P/Invoke), access unmanaged code.
    ''' This class does not suppress stack walks for unmanaged code permission.
    ''' <see cref="System.Security.SuppressUnmanagedCodeSecurityAttribute"/> must not be applied to this class.
    ''' This class is for methods that can be used anywhere because a stack walk will be performed.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="http://msdn.microsoft.com/en-us/library/ms182161.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    Private NotInheritable Class NativeMethods

#Region " Functions "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sends the specified message to a window or windows.
        ''' The SendMessage function calls the window procedure for the specified window
        ''' and does not return until the window procedure has processed the message.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' A handle to the window whose window procedure will receive the message.
        ''' </param>
        ''' 
        ''' <param name="msg">
        ''' The message to be sent.
        ''' </param>
        ''' 
        ''' <param name="wParam">
        ''' Additional message-specific information.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' Additional message-specific information.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The return value specifies the result of the message processing; it depends on the message sent.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644950%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SendMessage(ByVal hWnd As IntPtr,
                                           ByVal msg As WindowsMessages,
                                           ByVal wParam As IntPtr,
                                           ByVal lParam As String
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sends the specified message to a window or windows.
        ''' The SendMessage function calls the window procedure for the specified window
        ''' and does not return until the window procedure has processed the message.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' A handle to the window whose window procedure will receive the message.
        ''' </param>
        ''' 
        ''' <param name="msg">
        ''' The message to be sent.
        ''' </param>
        ''' 
        ''' <param name="wParam">
        ''' Additional message-specific information.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' Additional message-specific information.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The return value specifies the result of the message processing; it depends on the message sent.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644950%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SendMessage(ByVal hWnd As IntPtr,
                                           ByVal msg As WindowsMessages,
                                           ByVal wParam As IntPtr,
                                           ByVal lParam As IntPtr
        ) As IntPtr
        End Function

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The system sends or posts a system-defined message when it communicates with an application. 
        ''' It uses these messages to control the operations of applications and to provide input and other information for applications to process. 
        ''' An application can also send or post system-defined messages.
        ''' Applications generally use these messages to control the operation of control windows created by using preregistered window classes.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644927%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Enum WindowsMessages As Integer

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Determines the length, in characters, of the text associated with a window.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
            ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms632628%28v=vs.85%29.aspx"/>
            ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            WmGetTextLength = &HE

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Sets the text of a window.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
            ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms632644%28v=vs.85%29.aspx"/>
            ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            WmSetText = &HC

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Selects a range of characters in an edit control. 
            ''' You can send this message to either an edit control or a rich edit control.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
            ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb761661%28v=vs.85%29.aspx"/>
            ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            EmcSetSel = &HB1

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Replaces the selected text in an edit control or a rich edit control with the specified text.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
            ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb761633%28v=vs.85%29.aspx"/>
            ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            EmReplaceSel = &HC2

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Specifies additional message-specific information for a System-Defined Message.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644927%28v=vs.85%29.aspx#system_defined"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Enum WParams As UInteger

            ''' <summary>
            ''' A Null WParam.
            ''' </summary>
            None = 0UI

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Specifies additional message-specific information for a System-Defined Message.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644927%28v=vs.85%29.aspx#system_defined"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Enum LParams As UInteger

            ''' <summary>
            ''' A Null LParam.
            ''' </summary>
            None = 0UI

        End Enum

#End Region

    End Class

#End Region

#Region " Child Classes "

#Region " Shared Memory "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related memory-sharing utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class SharedMemory

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates a <see cref="MemoryMappedFile"/> segment that is shared between applications.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name to assign the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="capacity">
        ''' The maximum size, in bytes, to allocate data on the <see cref="MemoryMappedFile"/> segment.
        ''' 
        ''' The specified value is automatically rounded to a multiple of 4096 bytes (4 KB),
        ''' for example a value of 1 will be rounded to 4096, a value of 4097 will be rounded to 8192, and a value of 9999 to 12288.
        ''' </param>
        ''' 
        ''' <param name="fileAccess">
        ''' The <see cref="MemoryMappedFileAccess"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function Create(ByVal name As String,
                                      ByVal capacity As Integer,
                                      Optional ByVal fileAccess As MemoryMappedFileAccess =
                                                                   MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile

            Return MemoryMappedFile.CreateNew(name, capacity, fileAccess)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Clears the data of an existing <see cref="MemoryMappedFile"/> segment.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Clear(mmf)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub Clear(ByVal mmf As MemoryMappedFile)

            Using writer As MemoryMappedViewAccessor = mmf.CreateViewAccessor()

                writer.WriteArray(Of Byte)(0L, Enumerable.Repeat(New Byte, CInt(writer.Capacity)).ToArray, 0I, CInt(writer.Capacity))

            End Using

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Clears the data of an existing <see cref="MemoryMappedFile"/> segment.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Clear("My MemoryMappedFile Name")
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub Clear(ByVal name As String)

            IpcUtil.SharedMemory.Clear(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a byte sequence at the start position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write(mmf, Encoding.Default.GetBytes("Hello World!"))
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="data">
        ''' The byte sequence to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub Write(ByVal mmf As MemoryMappedFile,
                                ByVal data As Byte())

            Using writer As MemoryMappedViewAccessor = mmf.CreateViewAccessor()

                writer.WriteArray(Of Byte)(0L, data, 0I, data.Length)

            End Using

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a byte sequence at the start position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write("My MemoryMappedFile Name", Encoding.Default.GetBytes("Hello World!"))
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="data">
        ''' The byte sequence to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub Write(ByVal name As String,
                                ByVal data As Byte())

            IpcUtil.SharedMemory.Write(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite), data)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a single byte at the specified position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteAt(mmf, New Byte, offset:=0)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The byte value to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="offset">
        ''' The position in the <see cref="MemoryMappedFile"/> segment to start writing from.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteAt(ByVal mmf As MemoryMappedFile,
                                  ByVal value As Byte,
                                  ByVal offset As Long)

            Using writer As MemoryMappedViewAccessor = mmf.CreateViewAccessor()

                writer.Write(Of Byte)(offset, value)

            End Using

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a single byte at the specified position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteAt("My MemoryMappedFile Name", New Byte, offset:=0)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The byte value to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="offset">
        ''' The position in the <see cref="MemoryMappedFile"/> segment to start writing from.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteAt(ByVal name As String,
                                  ByVal value As Byte,
                                  ByVal offset As Long)

            IpcUtil.SharedMemory.WriteAt(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite), value, offset)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a byte sequence at the specified position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteAt(mmf, Encoding.Default.GetBytes("Hello World!"), offset:=0)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="data">
        ''' The byte sequence to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="offset">
        ''' The position in the <see cref="MemoryMappedFile"/> segment to start writing from.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteAt(ByVal mmf As MemoryMappedFile,
                                  ByVal data As Byte(),
                                  ByVal offset As Long)

            Using writer As MemoryMappedViewAccessor = mmf.CreateViewAccessor()

                writer.WriteArray(Of Byte)(offset, data, 0I, data.Length)

            End Using

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a byte sequence at the specified position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteAt("My MemoryMappedFile Name", Encoding.Default.GetBytes("Hello World!"), offset:=0)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="data">
        ''' The byte sequence to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="offset">
        ''' The position in the <see cref="MemoryMappedFile"/> segment to start writing from.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteAt(ByVal name As String,
                                  ByVal data As Byte(),
                                  ByVal offset As Long)

            IpcUtil.SharedMemory.WriteAt(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite), data, offset)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads from start to end the data of an existing <see cref="MemoryMappedFile"/>.
        '''
        ''' Note that the returned bytes could contain null bytes at the end 
        ''' due to the automatic size rounding of a multiple of 4096 bytes (4 KB).
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim enc As Encoding = Encoding.Default
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write(mmf, enc.GetBytes(str))
        '''     result = enc.GetString(IpcUtil.SharedMemory.ReadToEnd(mmf)).TrimEnd({ControlChars.NullChar})
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The byte-data.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadToEnd(ByVal mmf As MemoryMappedFile) As Byte()

            Using stream As MemoryMappedViewStream = mmf.CreateViewStream()

                stream.Seek(0, SeekOrigin.Begin)

                Using reader As New BinaryReader(stream)

                    Return reader.ReadBytes(CInt(stream.Length))

                End Using ' reader

            End Using ' stream

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads from start to end the data of an existing <see cref="MemoryMappedFile"/>.
        '''
        ''' Note that the returned bytes could contain null bytes at the end 
        ''' due to the automatic size rounding of a multiple of 4096 bytes (4 KB).
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim enc As Encoding = Encoding.Default
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write("My MemoryMappedFile Name", enc.GetBytes(str))
        '''     result = enc.GetString(IpcUtil.SharedMemory.ReadToEnd("My MemoryMappedFile Name")).TrimEnd({ControlChars.NullChar})
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The byte-data.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadToEnd(ByVal name As String) As Byte()

            Return IpcUtil.SharedMemory.ReadToEnd(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite))

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads a byte from a position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write(mmf, Encoding.Default.GetBytes(str))
        '''     result = Convert.ToChar(IpcUtil.SharedMemory.ReadAt(mmf, offset:=0))
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The byte value.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadAt(ByVal mmf As MemoryMappedFile, ByVal offset As Long) As Byte

            Dim buffer As Byte() = New Byte(1) {}

            Using reader As MemoryMappedViewStream = mmf.CreateViewStream()

                reader.Seek(offset, SeekOrigin.Begin)
                reader.Read(buffer, 0L, 1)
                Return buffer.First

            End Using

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads a byte from a position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write("My MemoryMappedFile Name", Encoding.Default.GetBytes(str))
        '''     result = Convert.ToChar(IpcUtil.SharedMemory.ReadAt("My MemoryMappedFile Name", offset:=0))
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The byte value.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadAt(ByVal name As String, ByVal offset As Long) As Byte

            Return IpcUtil.SharedMemory.ReadAt(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite), offset)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads a byte sequence from a start position to an end position of an existing <see cref="MemoryMappedFile"/>.
        '''
        ''' Note that the returned bytes could contain null bytes at the end 
        ''' due to the automatic size rounding of a multiple of 4096 bytes (4 KB).
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim enc As Encoding = Encoding.Default
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write(mmf, enc.GetBytes(str))
        '''     result = enc.GetString(IpcUtil.SharedMemory.Read(mmf, starIndex:=0, endIndex:=5))
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="starIndex">
        ''' The start position to start reading from the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="endIndex">
        ''' The end position to stop reading from the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The byte sequence.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function Read(ByVal mmf As MemoryMappedFile,
                                    ByVal starIndex As Long,
                                    ByVal endIndex As Long) As Byte()

            Using stream As MemoryMappedViewStream = mmf.CreateViewStream()

                stream.Seek(starIndex, SeekOrigin.Begin)

                Dim length As Integer = CInt((endIndex - starIndex) + stream.Position)
                Dim buffer As Byte() = New Byte(length) {}

                stream.Read(buffer, 0, length)
                Return buffer

            End Using

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads a byte sequence from a start position to an end position of an existing <see cref="MemoryMappedFile"/>.
        '''
        ''' Note that the returned bytes could contain null bytes at the end 
        ''' due to the automatic size rounding of a multiple of 4096 bytes (4 KB).
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim enc As Encoding = Encoding.Default
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write("My MemoryMappedFile Name", enc.GetBytes(str))
        '''     result = enc.GetString(IpcUtil.SharedMemory.Read("My MemoryMappedFile Name", starIndex:=0, endIndex:=5))
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name o the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="starIndex">
        ''' The start position to start reading from the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="endIndex">
        ''' The end position to stop reading from the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The byte sequence.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function Read(ByVal name As String,
                                    ByVal starIndex As Long,
                                    ByVal endIndex As Long) As Byte()

            Return IpcUtil.SharedMemory.Read(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite), starIndex, endIndex)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a String at the start position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteString(mmf, "Hello World!", Encoding.Default)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="str">
        ''' The text to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The text <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteString(ByVal mmf As MemoryMappedFile,
                                      ByVal str As String,
                                      Optional ByVal encoding As Encoding = Nothing)

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            IpcUtil.SharedMemory.Write(mmf, encoding.GetBytes(str))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a String at the start position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteString("My MemoryMappedFile Name", "Hello World!", Encoding.Default)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="str">
        ''' The text to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The text <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteString(ByVal name As String,
                                      ByVal str As String,
                                      Optional ByVal encoding As Encoding = Nothing)

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            IpcUtil.SharedMemory.Write(name, encoding.GetBytes(str))

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a character at the specified position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteCharAt(mmf, "A"c, offset:=0, encoding:=Encoding.Default)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="c">
        ''' The character to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="offset">
        ''' The position in the <see cref="MemoryMappedFile"/> segment to start writing from.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The character <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteCharAt(ByVal mmf As MemoryMappedFile,
                                      ByVal c As Char,
                                      ByVal offset As Long,
                                      Optional ByVal encoding As Encoding = Nothing)

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            IpcUtil.SharedMemory.WriteAt(mmf, encoding.GetBytes(c), offset)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a character at the specified position of an existing <see cref="MemoryMappedFile"/> segment.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteStringAt("My MemoryMappedFile Name","A"c, offset:=0, encoding:=Encoding.Default)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="c">
        ''' The character to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="offset">
        ''' The position in the <see cref="MemoryMappedFile"/> segment to start writing from.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The character <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteCharAt(ByVal name As String,
                                      ByVal c As Char,
                                      ByVal offset As Long,
                                      Optional ByVal encoding As Encoding = Nothing)

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            IpcUtil.SharedMemory.WriteAt(name, encoding.GetBytes(c), offset)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a String at the specified position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteStringAt(mmf, "Hello World!", offset:=0, encoding:=Encoding.Default)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="str">
        ''' The text to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="offset">
        ''' The position in the <see cref="MemoryMappedFile"/> segment to start writing from.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The text <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteStringAt(ByVal mmf As MemoryMappedFile,
                                        ByVal str As String,
                                        ByVal offset As Long,
                                        Optional ByVal encoding As Encoding = Nothing)

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            IpcUtil.SharedMemory.WriteAt(mmf, encoding.GetBytes(str), offset)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Writes a String at the specified position of an existing <see cref="MemoryMappedFile"/> segment.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.WriteStringAt("My MemoryMappedFile Name", "Hello World!", offset:=0, encoding:=Encoding.Default)
        ''' 
        ''' End Using
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="str">
        ''' The text to write in the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="offset">
        ''' The position in the <see cref="MemoryMappedFile"/> segment to start writing from.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The text <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub WriteStringAt(ByVal name As String,
                                        ByVal str As String,
                                        ByVal offset As Long,
                                        Optional ByVal encoding As Encoding = Nothing)

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            IpcUtil.SharedMemory.WriteAt(name, encoding.GetBytes(str), offset)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads from start to end the data of an existing <see cref="MemoryMappedFile"/>,
        ''' decodes the byte data using the specified <see cref="Encoding"/> and returns the corresponding string.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim enc As Encoding = Encoding.Default
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write(mmf, enc.GetBytes(str))
        '''     result = IpcUtil.SharedMemory.ReadStringToEnd(mmf, enc)
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The text <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The byte-data.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadStringToEnd(ByVal mmf As MemoryMappedFile,
                                               Optional ByVal encoding As Encoding = Nothing) As String

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            Return encoding.GetString(IpcUtil.SharedMemory.ReadToEnd(mmf)).
                            Trim({ControlChars.NullChar})

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads from start to end the data of an existing <see cref="MemoryMappedFile"/>,
        ''' decodes the byte data using the specified <see cref="Encoding"/> and returns the corresponding string.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim enc As Encoding = Encoding.Default
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write("My MemoryMappedFile Name", enc.GetBytes(str))
        '''     result = IpcUtil.SharedMemory.ReadStringToEnd("My MemoryMappedFile Name", enc)
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The text <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The byte-data.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadStringToEnd(ByVal name As String,
                                               Optional ByVal encoding As Encoding = Nothing) As String

            Return IpcUtil.SharedMemory.ReadStringToEnd(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite), encoding)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads a character from a position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim str As String = "Hello World!"
        ''' Dim result As Char
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write(mmf, Encoding.Default.GetBytes(str))
        '''     result = IpcUtil.SharedMemory.ReadCharAt(mmf, offset:=0)
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The character.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadCharAt(ByVal mmf As MemoryMappedFile, ByVal offset As Long) As Char

            Return Convert.ToChar(IpcUtil.SharedMemory.ReadAt(mmf, offset))

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads a character from a position of an existing <see cref="MemoryMappedFile"/>.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim str As String = "Hello World!"
        ''' Dim result As Char
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write("My MemoryMappedFile Name", Encoding.Default.GetBytes(str))
        '''     result = IpcUtil.SharedMemory.ReadCharAt("My MemoryMappedFile Name", offset:=0)
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The character.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadCharAt(ByVal name As String, ByVal offset As Long) As Char

            Return Convert.ToChar(IpcUtil.SharedMemory.ReadAt(MemoryMappedFile.OpenExisting(name, MemoryMappedFileRights.ReadWrite), offset))

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads a byte sequence from a start position to an end position of an existing <see cref="MemoryMappedFile"/>,
        ''' decodes the byte data using the specified <see cref="Encoding"/> and returns the corresponding string.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim enc As Encoding = Encoding.Default
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write(mmf, enc.GetBytes(str))
        '''     result = IpcUtil.SharedMemory.ReadString(mmf, startIndex:=0, endIndex:=5)
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="mmf">
        ''' The <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="startIndex">
        ''' The start position to start reading from the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="endIndex">
        ''' The end position to stop reading from the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The text <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The decoded string.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadString(ByVal mmf As MemoryMappedFile,
                                          ByVal startIndex As Long,
                                          ByVal endIndex As Long,
                                          Optional ByVal encoding As Encoding = Nothing) As String

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            Return encoding.GetString(IpcUtil.SharedMemory.Read(mmf, startIndex, endIndex)).
                            Trim({ControlChars.NullChar})

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reads a byte sequence from a start position to an end position of an existing <see cref="MemoryMappedFile"/>,
        ''' decodes the byte data using the specified <see cref="Encoding"/> and returns the corresponding string.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim enc As Encoding = Encoding.Default
        ''' Dim str As String = "Hello World!"
        ''' Dim result As String
        ''' 
        ''' Using mmf As MemoryMappedFile = IpcUtil.SharedMemory.Create("My MemoryMappedFile Name", capacity:=4096)
        ''' 
        '''     IpcUtil.SharedMemory.Write("My MemoryMappedFile Name", enc.GetBytes(str))
        '''     result = IpcUtil.SharedMemory.ReadString("My MemoryMappedFile Name", startIndex:=0, endIndex:=5)
        ''' 
        ''' End Using
        ''' 
        ''' MessageBox.Show(result)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="name">
        ''' The name of the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="starIndex">
        ''' The start position to start reading from the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="endIndex">
        ''' The end position to stop reading from the <see cref="MemoryMappedFile"/> segment.
        ''' </param>
        ''' 
        ''' <param name="encoding">
        ''' The text <see cref="Encoding"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The decoded string.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ReadString(ByVal name As String,
                                          ByVal starIndex As Long,
                                          ByVal endIndex As Long,
                                          Optional ByVal encoding As Encoding = Nothing) As String

            If (encoding Is Nothing) Then
                encoding = System.Text.Encoding.ASCII
            End If

            Return encoding.GetString(IpcUtil.SharedMemory.Read(name, starIndex, endIndex)).
                            Trim({ControlChars.NullChar})

        End Function

#End Region

    End Class

#End Region

#Region " UI Automation "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related UI automation utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class UIAutomation

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the titlebar's text of the specified window.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' MessageBox.Show(IpcUtil.GetTitlebarText(Process.GetCurrentProcess.MainWindowHandle))
        ''' MessageBox.Show(IpcUtil.GetTitlebarText(Process.GetProcessesByName("notepad").First.MainWindowHandle))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' An <see cref="IntPtr"/> to the window.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The titlebar's text of the target window.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function GetTitlebarText(ByVal hWnd As IntPtr) As String

            Dim window As AutomationElement = AutomationElement.FromHandle(hWnd)
            Dim condition As New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TitleBar)
            Dim titleBar As AutomationElement = window.FindFirst(TreeScope.Children, condition)

            Return titleBar.Current.Name

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the text of an Edit control.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwnd">
        ''' A <see cref="IntPtr"/> handle to the Edit window.
        ''' </param>
        ''' 
        ''' <param name="text">
        ''' The text.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if operation succeeds, <see langword="True"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Function SetText(ByVal hwnd As IntPtr,
                                ByVal text As String) As Boolean

            Return CBool(NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.WmSetText,
                                                   New IntPtr(NativeMethods.WParams.None), text))

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Appends text at the end of an Edit control.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwnd">
        ''' A <see cref="IntPtr"/> handle to the Edit window.
        ''' </param>
        ''' 
        ''' <param name="text">
        ''' The text to append.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if operation succeeds, <see langword="True"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Function AppendText(ByVal hwnd As IntPtr,
                                   ByVal text As String) As Boolean

            ' Get text length.
            Dim textLength As Integer =
                NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.WmGetTextLength,
                                          New IntPtr(NativeMethods.WParams.None),
                                          New IntPtr(NativeMethods.LParams.None)).ToInt32

            ' Set text selection.
            NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.EmcSetSel,
                                      New IntPtr(textLength), New IntPtr(-1))

            ' Replace selected text.
            Return CBool(NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.EmReplaceSel,
                                                   New IntPtr(1), text))

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Inserts text at the specified position of an Edit control.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwnd">
        ''' A <see cref="IntPtr"/> handle to the Edit window.
        ''' </param>
        ''' 
        ''' <param name="position">
        ''' The character position where to insert the text.
        ''' </param>
        ''' 
        ''' <param name="text">
        ''' The text to insert.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if operation succeeds, <see langword="True"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        Public Function InsertText(ByVal hwnd As IntPtr,
                                   ByVal position As Integer,
                                   ByVal text As String) As Boolean

            ' Set text selection.
            NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.EmcSetSel,
                                      New IntPtr(position), New IntPtr(position))

            ' Replace selected text.
            Return CBool(NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.EmReplaceSel,
                                                   New IntPtr(1), Text))

        End Function

#End Region

    End Class

#End Region

#End Region

End Module

#End Region
