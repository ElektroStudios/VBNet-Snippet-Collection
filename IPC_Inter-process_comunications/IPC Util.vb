' ***********************************************************************
' Author   : Elektro
' Modified : 04-November-2015
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

' IpcUtil.UIAutomation.AppendText(IntPtr, String) As Boolean
' IpcUtil.UIAutomation.GetText(IntPtr) As String
' IpcUtil.UIAutomation.GetTitlebarText(IntPtr) As String
' IpcUtil.UIAutomation.InsertText(IntPtr, Integer, String) As Boolean
' IpcUtil.UIAutomation.MoveWindow(Process, Point, Opt: Boolean) As Boolean
' IpcUtil.UIAutomation.MoveWindow(String, Point, Opt: Boolean) As Boolean
' IpcUtil.UIAutomation.ResizeWindow(Process, Size, Opt: Boolean) As Boolean
' IpcUtil.UIAutomation.ResizeWindow(String, Size, Opt: Boolean) As Boolean
' IpcUtil.UIAutomation.SetText(IntPtr, String) As Boolean
' IpcUtil.UIAutomation.SetWindowState(IntPtr, ProcessUtil.WindowState) As Boolean
' IpcUtil.UIAutomation.SliceWindowPosition(Process, Point, Opt: Boolean) As Boolean
' IpcUtil.UIAutomation.SliceWindowPosition(String, Point, Opt: Boolean) As Boolean
' IpcUtil.UIAutomation.SliceWindowSize(Process, Size, Opt: Boolean) As Boolean
' IpcUtil.UIAutomation.SliceWindowSize(String, Size, Opt: Boolean) As Boolean

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
        ''' Retrieves the dimensions of the bounding rectangle of the specified window. 
        ''' The dimensions are given in screen coordinates that are relative to the upper-left corner of the screen.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' A handle to the window.
        ''' </param>
        ''' 
        ''' <param name="rect">
        ''' A pointer to a <see cref="IpcUtil.NativeMethods.Rect"/> structure that receives the screen coordinates of the 
        ''' upper-left and lower-right corners of the window.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if If the function succeeds, <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/es-es/library/windows/desktop/ms633519%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function GetWindowRect(ByVal hWnd As IntPtr,
                                             ByRef rect As IpcUtil.NativeMethods.Rect
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Changes the position and dimensions of the specified window. 
        ''' For a top-level window, the position and dimensions are relative to the upper-left corner of the screen. 
        ''' For a child window, they are relative to the upper-left corner of the parent window's client area.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' A handle to the window.
        ''' </param>
        ''' 
        ''' <param name="x">
        ''' The new position of the left side of the window.
        ''' </param>
        ''' 
        ''' <param name="y">
        ''' The new position of the top of the window.
        ''' </param>
        ''' 
        ''' <param name="width">
        ''' The new width of the window.
        ''' </param>
        ''' 
        ''' <param name="height">
        ''' The new height of the window.
        ''' </param>
        ''' 
        ''' <param name="repaint">
        ''' Indicates whether the window is to be repainted. 
        ''' If this parameter is <see langword="True"/>, the window receives a message. 
        ''' If the parameter is <see langword="False"/>, no repainting of any kind occurs. 
        ''' This applies to the client area, the nonclient area (including the title bar and scroll bars), 
        ''' and any part of the parent window uncovered as a result of moving a child window.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if If the function succeeds, <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms633534%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function MoveWindow(ByVal hWnd As IntPtr,
                                          ByVal x As Integer,
                                          ByVal y As Integer,
                                          ByVal width As Integer,
                                          ByVal height As Integer,
          <MarshalAs(UnmanagedType.Bool)> ByVal repaint As Boolean
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves a handle to the top-level window whose class name and window name match the specified strings.
        ''' This function does not search child windows.
        ''' This function does not perform a case-sensitive search.
        ''' To search child windows, beginning with a specified child window, use the <see cref="FindWindowEx"/> function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="lpClassName">
        ''' The class name.
        ''' If this parameter is <see langword="Nothing"/>, 
        ''' it finds any window whose title matches the <paramref name="lpWindowName"/> parameter.
        ''' </param>
        ''' 
        ''' <param name="lpWindowName">
        ''' The window name (the window title).
        ''' If this parameter is <see langword="Nothing"/>, all window names match.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a handle to the window that has the specified class name and window name.
        ''' If the function fails, the return value is <see langword="IntPtr.Zero"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms633499%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappablechar:=True)>
        Friend Shared Function FindWindow(ByVal lpClassName As String,
                                          ByVal lpWindowName As String
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves a handle to a window whose class name and window name match the specified strings. 
        ''' The function searches child windows, beginning with the one following the specified child window. 
        ''' This function does not perform a case-sensitive search.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwndParent">
        ''' A <see cref="IntPtr"/> handle to the parent window whose child windows are to be searched.
        ''' If <paramref name="hwndParent"/> is <see cref="IntPtr.Zero"/>, the function uses the desktop window as the parent window. 
        ''' The function searches among windows that are child windows of the desktop. 
        ''' </param>
        ''' 
        ''' <param name="hwndChildAfter">
        ''' A <see cref="IntPtr"/> handle to a child window. 
        ''' The search begins with the next child window in the Z order. 
        ''' The child window must be a direct child window of hwndParent, not just a descendant window.
        ''' If <paramref name="hwndChildAfter"/> is <see cref="IntPtr.Zero"/>, 
        ''' the search begins with the first child window of <paramref name="hwndParent"/>.
        ''' </param>
        ''' 
        ''' <param name="strClassName">
        ''' The window class name.
        ''' </param>
        ''' 
        ''' <param name="strWindowName">
        ''' The window name (the window title). 
        ''' If this parameter is <see langword="Nothing"/>, all window names match.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a <see cref="IntPtr"/> handle to the window that has the specified class and window names.
        ''' If the function fails, the return value is NULL.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms633500%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappablechar:=True)>
        Friend Shared Function FindWindowEx(ByVal hwndParent As IntPtr,
                                            ByVal hwndChildAfter As IntPtr,
                                            ByVal strClassName As String,
                                            ByVal strWindowName As String
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the identifier of the thread that created the specified window 
        ''' and, optionally, the identifier of the process that created the window.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' A <see cref="IntPtr"/> handle to the window.
        ''' </param>
        ''' 
        ''' <param name="processId">
        ''' A pointer to a variable that receives the process identifier. 
        ''' If this parameter is not <see langword="Nothing"/>, <see cref="GetWindowThreadProcessId"/> copies the identifier of 
        ''' the process to the variable; otherwise, it does not.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The identifier of the thread that created the window.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms633522%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll")>
        Friend Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr,
                                                        ByRef processId As Integer
        ) As Integer
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the specified window's show state.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwnd">
        ''' A <see cref="IntPtr"/> handle to the window.
        ''' </param>
        ''' 
        ''' <param name="nCmdShow">
        ''' Controls how the window is to be shown.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the window was previously visible, the return value is <see langword="True"/>. 
        ''' If the window was previously hidden, the return value is <see langword="False"/>. 
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("User32", SetLastError:=False)>
        Friend Shared Function ShowWindow(ByVal hwnd As IntPtr,
                                          ByVal nCmdShow As IpcUtil.UIAutomation.WindowState
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
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
                                           ByVal lParam As StringBuilder
        ) As IntPtr
        End Function

        ' ''' ----------------------------------------------------------------------------------------------------
        ' ''' <summary>
        ' ''' Reads data from an area of memory in a specified process. 
        ' ''' The entire area to be read must be accessible or the operation fails.
        ' ''' </summary>
        ' ''' ----------------------------------------------------------------------------------------------------
        ' ''' <param name="hProcess">
        ' ''' A handle to the process with memory that is being read. 
        ' ''' The handle must have PROCESS_VM_READ access to the process.
        ' ''' </param>
        ' ''' 
        ' ''' <param name="lpBaseAddress">
        ' ''' A pointer to the base address in the specified process from which to read. 
        ' ''' Before any data transfer occurs, the system verifies that all data in the base address and memory of the 
        ' ''' specified size is accessible for read access, and if it is not accessible the function fails.
        ' ''' </param>
        ' ''' 
        ' ''' <param name="lpBuffer">
        ' ''' A pointer to a buffer that receives the contents from the address space of the specified process.
        ' ''' </param>
        ' ''' 
        ' ''' <param name="iSize">
        ' ''' The number of bytes to be read from the specified process.
        ' ''' </param>
        ' ''' 
        ' ''' <param name="lpNumberOfBytesRead">
        ' ''' A pointer to a variable that receives the number of bytes transferred into the specified buffer. 
        ' ''' If <paramref name="lpNumberOfBytesRead"/> is <c>0</c>, the parameter is ignored.
        ' ''' </param>
        ' ''' ----------------------------------------------------------------------------------------------------
        ' ''' <returns>
        ' ''' <see langword="True"/> If the function succeeds, <see langword="False"/> otherwise.
        ' ''' </returns>
        ' ''' ----------------------------------------------------------------------------------------------------
        '<DllImport("kernel32.dll", SetLastError:=True)>
        'Friend Shared Function ReadProcessMemory(ByVal hProcess As IntPtr,
        '                                         ByVal lpBaseAddress As IntPtr,
        '                                         ByVal lpBuffer As IntPtr,
        '                                         ByVal iSize As Integer,
        '                                         ByRef lpNumberOfBytesRead As Integer
        ') As <MarshalAs(UnmanagedType.Bool)> Boolean
        'End Function

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
            ''' Copies the text that corresponds to a window into a buffer provided by the caller.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
            ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms632627%28v=vs.85%29.aspx"/>
            ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            WmGetText = &HD

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

#Region " Structures "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Defines the coordinates of the upper-left and lower-right corners of a rectangle.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/dd162897%28v=vs.85%29.aspx"/>
        ''' <para></para>
        ''' <see href="http://www.pinvoke.net/default.aspx/Structures/rect.html"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <StructLayout(LayoutKind.Sequential)>
        Friend Structure Rect

#Region " Properties "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets or sets the x-coordinate of the upper-left corner of the rectangle.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The x-coordinate of the upper-left corner of the rectangle.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public Property Left As Integer

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets or sets the y-coordinate of the upper-left corner of the rectangle.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The y-coordinate of the upper-left corner of the rectangle.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public Property Top As Integer

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets or sets the x-coordinate of the lower-right corner of the rectangle.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The x-coordinate of the lower-right corner of the rectangle.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public Property Right As Integer

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets or sets the y-coordinate of the lower-right corner of the rectangle.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The y-coordinate of the lower-right corner of the rectangle.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public Property Bottom As Integer

#End Region

#Region " Constructors "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Initializes a new instance of the <see cref="IpcUtil.NativeMethods.Rect"/> struct.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <param name="left">
            ''' The x-coordinate of the upper-left corner of the rectangle.
            ''' </param>
            ''' 
            ''' <param name="top">
            ''' The y-coordinate of the upper-left corner of the rectangle.
            ''' </param>
            ''' 
            ''' <param name="right">
            ''' The x-coordinate of the lower-right corner of the rectangle.
            ''' </param>
            ''' 
            ''' <param name="bottom">
            ''' The y-coordinate of the lower-right corner of the rectangle.
            ''' </param>
            ''' ----------------------------------------------------------------------------------------------------
            Public Sub New(ByVal left As Integer,
                           ByVal top As Integer,
                           ByVal right As Integer,
                           ByVal bottom As Integer)

                Me.Left = left
                Me.Top = top
                Me.Right = right
                Me.Bottom = bottom

            End Sub

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Initializes a new instance of the <see cref="NativeMethods.Rect"/> struct.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <param name="rect">
            ''' The <see cref="Rectangle"/>.
            ''' </param>
            ''' ----------------------------------------------------------------------------------------------------
            Public Sub New(ByVal rect As Rectangle)

                Me.New(rect.Left, rect.Top, rect.Right, rect.Bottom)

            End Sub

#End Region

#Region " Operator conversions "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Performs an implicit conversion from <see cref="NativeMethods.Rect"/> to <see cref="Rectangle"/>.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <param name="rect">The <see cref="NativeMethods.Rect"/>.
            ''' </param>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <returns>
            ''' The resulting <see cref="Rectangle"/>.
            ''' </returns>
            ''' ----------------------------------------------------------------------------------------------------
            Public Shared Widening Operator CType(rect As IpcUtil.NativeMethods.Rect) As Rectangle

                Return New Rectangle(rect.Left, rect.Top, (rect.Right - rect.Left), (rect.Bottom - rect.Top))

            End Operator

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Performs an implicit conversion from <see cref="Rectangle"/> to <see cref="NativeMethods.Rect"/>.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <param name="rect">The <see cref="Rectangle"/>.
            ''' </param>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <returns>
            ''' The resulting <see cref="NativeMethods.Rect"/>.
            ''' </returns>
            ''' ----------------------------------------------------------------------------------------------------
            Public Shared Widening Operator CType(rect As Rectangle) As IpcUtil.NativeMethods.Rect

                Return New IpcUtil.NativeMethods.Rect(rect)

            End Operator

#End Region

        End Structure

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

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Controls how a window is to be shown.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Public Enum WindowState As Integer

            ''' <summary>
            ''' Hides the window and activates another window.
            ''' </summary>
            Hide = 0

            ''' <summary>
            ''' Activates and displays a window. 
            ''' If the window is minimized or maximized, the system restores it to its original size and position.
            ''' An application should specify this flag when displaying the window for the first time.
            ''' </summary>
            Normal = 1

            ''' <summary>
            ''' Activates the window and displays it as a minimized window.
            ''' </summary>
            ShowMinimized = 2

            ''' <summary>
            ''' Maximizes the specified window.
            ''' </summary>
            Maximize = 3

            ''' <summary>
            ''' Activates the window and displays it as a maximized window.
            ''' </summary>      
            ShowMaximized = Maximize

            ''' <summary>
            ''' Displays a window in its most recent size and position. 
            ''' This value is similar to <see cref="WindowState.Normal"/>, except the window is not actived.
            ''' </summary>
            ShowNoActivate = 4

            ''' <summary>
            ''' Activates the window and displays it in its current size and position.
            ''' </summary>
            Show = 5

            ''' <summary>
            ''' Minimizes the specified window and activates the next top-level window in the Z order.
            ''' </summary>
            Minimize = 6

            ''' <summary>
            ''' Displays the window as a minimized window. 
            ''' This value is similar to <see cref="WindowState.ShowMinimized"/>, except the window is not activated.
            ''' </summary>
            ShowMinNoActive = 7

            ''' <summary>
            ''' Displays the window in its current size and position.
            ''' This value is similar to <see cref="WindowState.Show"/>, except the window is not activated.
            ''' </summary>
            ShowNA = 8

            ''' <summary>
            ''' Activates and displays the window. 
            ''' If the window is minimized or maximized, the system restores it to its original size and position.
            ''' An application should specify this flag when restoring a minimized window.
            ''' </summary>
            Restore = 9

            ''' <summary>
            ''' Sets the show state based on the SW_* value specified in the STARTUPINFO structure 
            ''' passed to the CreateProcess function by the program that started the application.
            ''' </summary>
            ShowDefault = 10

            ''' <summary>
            ''' <b>Windows 2000/XP:</b> 
            ''' Minimizes a window, even if the thread that owns the window is not responding. 
            ''' This flag should only be used when minimizing windows from a different thread.
            ''' </summary>
            ForceMinimize = 11

        End Enum

#End Region

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
        ''' Gets the text of an Edit control.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwnd">
        ''' A <see cref="IntPtr"/> handle to the Edit window.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The text.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Exception">
        ''' Invalid handle, window not found.
        ''' </exception>
        ''' 
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function GetText(ByVal hwnd As IntPtr) As String

            Dim win32Err As Integer
            Dim textLength As Integer =
                NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.WmGetTextLength,
                                          New IntPtr(NativeMethods.WParams.None),
                                          New IntPtr(NativeMethods.LParams.None)).ToInt32

            win32Err = Marshal.GetLastWin32Error
            If (win32Err = 1400) Then
                Throw New Exception(message:="Invalid handle, window not found.")

                'ElseIf (win32Err <> 1) Then
                '    Throw New Win32Exception([error]:=win32Err)

            Else
                Dim sb As New StringBuilder(capacity:=textLength)
                NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.WmGetText, New IntPtr(textLength), sb).ToInt32()
                Return sb.ToString

            End If

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
        ''' <exception cref="Exception">
        ''' Invalid handle, window not found.
        ''' </exception>
        ''' 
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function SetText(ByVal hwnd As IntPtr,
                                       ByVal text As String) As Boolean

            Dim result As Boolean
            Dim win32Err As Integer

            result = CBool(NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.WmSetText,
                                                     New IntPtr(NativeMethods.WParams.None), text))

            win32Err = Marshal.GetLastWin32Error
            If (win32Err = 1400) Then
                Throw New Exception(message:="Invalid handle, window not found.")

                'ElseIf (win32Err <> 1) Then
                '    Throw New Win32Exception([error]:=win32Err)

            Else
                Return result

            End If

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
        ''' <exception cref="Exception">
        ''' Invalid handle, window not found.
        ''' </exception>
        ''' 
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function AppendText(ByVal hwnd As IntPtr,
                                          ByVal text As String) As Boolean

            Dim win32Err As Integer
            Dim textLength As Integer =
                NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.WmGetTextLength,
                                          New IntPtr(NativeMethods.WParams.None),
                                          New IntPtr(NativeMethods.LParams.None)).ToInt32

            win32Err = Marshal.GetLastWin32Error
            If (win32Err = 1400) Then
                Throw New Exception(message:="Invalid handle, window not found.")

                'ElseIf (win32Err <> 1) Then
                '    Throw New Win32Exception([error]:=win32Err)

            Else
                ' Set text selection.
                NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.EmcSetSel,
                                          New IntPtr(textLength), New IntPtr(-1))

                ' Replace selected text.
                Return CBool(NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.EmReplaceSel,
                                                       New IntPtr(1), text))

            End If

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
        ''' <exception cref="Exception">
        ''' Invalid handle, window not found.
        ''' </exception>
        ''' 
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function InsertText(ByVal hwnd As IntPtr,
                                          ByVal position As Integer,
                                          ByVal text As String) As Boolean

            Dim win32Err As Integer

            ' Set text selection.
            NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.EmcSetSel,
                                      New IntPtr(position), New IntPtr(position))

            win32Err = Marshal.GetLastWin32Error
            If (win32Err = 1400) Then
                Throw New Exception(message:="Invalid handle, window not found.")

                'ElseIf (win32Err <> 1) Then
                '    Throw New Win32Exception([error]:=win32Err)

            Else
                ' Replace selected text.
                Return CBool(NativeMethods.SendMessage(hwnd, NativeMethods.WindowsMessages.EmReplaceSel,
                                                       New IntPtr(1), text))

            End If

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Moves the window of the first ocurrence found of a running process with the specified name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' IpcUtil.UIAutomation.MoveWindow("notepad.exe", New Point(0, 0))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="processName">
        ''' The process name.
        ''' </param>
        ''' 
        ''' <param name="position">
        ''' The new window position.
        ''' </param>
        ''' 
        ''' <param name="throwOnProcessNotFound">
        ''' If <see langword="True"/>, throws an <see cref="ArgumentException"/> exception if any process was found with the specified name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' if successful <see langword="True"/>; if failed <see langword="True"/> or
        ''' if <paramref name="throwOnProcessNotFound"/> is <see langword="False"/> and any process was found with the specified name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Any process found with the specified name.;processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function MoveWindow(ByVal processName As String,
                                          ByVal position As Point,
                                          Optional throwOnProcessNotFound As Boolean = False) As Boolean

            Return IpcUtil.UIAutomation.MoveWindow(Process.GetProcessesByName(IpcUtil.UIAutomation.FixProcessName(processName)).FirstOrDefault,
                                                   position, throwOnProcessNotFound)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Moves the window of the specified process.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' IpcUtil.UIAutomation.MoveWindow(Process.GetProcessesByName("notepad").FirstOrDefault, New Point(0, 0))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="p">
        ''' The process.
        ''' </param>
        ''' 
        ''' <param name="position">
        ''' The new window position.
        ''' </param>
        ''' 
        ''' <param name="throwOnProcessNotFound">
        ''' If <see langword="True"/>, throws an <see cref="ArgumentException"/> exception if any process was found with the specified name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' if successful <see langword="True"/>; if failed <see langword="True"/> or
        ''' if <paramref name="throwOnProcessNotFound"/> is <see langword="False"/> and any process was found with the specified name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Any process found with the specified name.;processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function MoveWindow(ByVal p As Process,
                                          ByVal position As Point,
                                          Optional throwOnProcessNotFound As Boolean = False) As Boolean

            If (p Is Nothing) AndAlso (throwOnProcessNotFound) Then
                Throw New ArgumentException(message:="Any process found with the specified name.", paramName:="processName")

            Else
                Using p

                    Dim rect As IpcUtil.NativeMethods.Rect ' Win32 Rectangle
                    IpcUtil.NativeMethods.GetWindowRect(p.MainWindowHandle, rect)

                    Dim rectangle As Rectangle = rect ' Managed Rectangle
                    Return IpcUtil.NativeMethods.MoveWindow(p.MainWindowHandle,
                                                            position.X, position.Y,
                                                            rectangle.Width, rectangle.Height, True)

                End Using

            End If

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Resizes the window of the first ocurrence found of a running process with the specified name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' IpcUtil.UIAutomation.SliceWindow("notepad.exe", New Size(640, 480))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="processName">
        ''' The process name.
        ''' </param>
        ''' 
        ''' <param name="size">
        ''' The new window size.
        ''' </param>
        ''' 
        ''' <param name="throwOnProcessNotFound">
        ''' If <see langword="True"/>, throws an <see cref="ArgumentException"/> exception if any process was found with the specified name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' if successful <see langword="True"/>; if failed <see langword="True"/> or
        ''' if <paramref name="throwOnProcessNotFound"/> is <see langword="False"/> and any process was found with the specified name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Any process found with the specified name.;processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ResizeWindow(ByVal processName As String,
                                            ByVal size As Size,
                                            Optional throwOnProcessNotFound As Boolean = False) As Boolean

            Return IpcUtil.UIAutomation.ResizeWindow(Process.GetProcessesByName(IpcUtil.UIAutomation.FixProcessName(processName)).FirstOrDefault,
                                                     size, throwOnProcessNotFound)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Resizes the window of the specified process.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' IpcUtil.UIAutomation.ResizeWindow(Process.GetProcessesByName("notepad").FirstOrDefault, New Size(640, 480))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="p">
        ''' The process.
        ''' </param>
        ''' 
        ''' <param name="size">
        ''' The new window size.
        ''' </param>
        ''' 
        ''' <param name="throwOnProcessNotFound">
        ''' If <see langword="True"/>, throws an <see cref="ArgumentException"/> exception if any process was found with the specified name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' if successful <see langword="True"/>; if failed <see langword="True"/> or
        ''' if <paramref name="throwOnProcessNotFound"/> is <see langword="False"/> and any process was found with the specified name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Any process found with the specified name.;processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ResizeWindow(ByVal p As Process,
                                            ByVal size As Size,
                                            Optional throwOnProcessNotFound As Boolean = False) As Boolean

            If (p Is Nothing) AndAlso (throwOnProcessNotFound) Then
                Throw New ArgumentException(message:="Any process found with the specified name.", paramName:="processName")

            Else
                Using p

                    Dim rect As IpcUtil.NativeMethods.Rect ' Win32 Rectangle
                    IpcUtil.NativeMethods.GetWindowRect(p.MainWindowHandle, rect)

                    Dim rectangle As Rectangle = rect ' Managed Rectangle
                    Return IpcUtil.NativeMethods.MoveWindow(p.MainWindowHandle,
                                                            rectangle.Left, rectangle.Top,
                                                            size.Width, size.Height, True)

                End Using

            End If

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Slices the window position of the first ocurrence found of a running process with the specified name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' IpcUtil.UIAutomation.SliceWindowPosition("notepad.exe", New Point(+10, -100))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="processName">
        ''' The process name.
        ''' </param>
        ''' 
        ''' <param name="position">
        ''' The new window position.
        ''' </param>
        ''' 
        ''' <param name="throwOnProcessNotFound">
        ''' If <see langword="True"/>, throws an <see cref="ArgumentException"/> exception if any process was found with the specified name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' if successful <see langword="True"/>; if failed <see langword="True"/> or
        ''' if <paramref name="throwOnProcessNotFound"/> is <see langword="False"/> and any process was found with the specified name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Any process found with the specified name.;processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function SliceWindowPosition(ByVal processName As String,
                                                   ByVal position As Point,
                                                   Optional throwOnProcessNotFound As Boolean = False) As Boolean

            Return IpcUtil.UIAutomation.SliceWindowPosition(Process.GetProcessesByName(IpcUtil.UIAutomation.FixProcessName(processName)).FirstOrDefault,
                                                            position, throwOnProcessNotFound)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Slices the window position of the specified process.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' IpcUtil.UIAutomation.SliceWindowPosition(Process.GetProcessesByName("notepad").FirstOrDefault, New Point(+10, -100))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="p">
        ''' The process.
        ''' </param>
        ''' 
        ''' <param name="position">
        ''' The new window position.
        ''' </param>
        ''' 
        ''' <param name="throwOnProcessNotFound">
        ''' If <see langword="True"/>, throws an <see cref="ArgumentException"/> exception if any process was found with the specified name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' if successful <see langword="True"/>; if failed <see langword="True"/> or
        ''' if <paramref name="throwOnProcessNotFound"/> is <see langword="False"/> and any process was found with the specified name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Any process found with the specified name.;processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function SliceWindowPosition(ByVal p As Process,
                                                   ByVal position As Point,
                                                   Optional throwOnProcessNotFound As Boolean = False) As Boolean

            If (p Is Nothing) AndAlso (throwOnProcessNotFound) Then
                Throw New ArgumentException(message:="Any process found with the specified name.", paramName:="processName")

            Else
                Using p

                    Dim rect As IpcUtil.NativeMethods.Rect ' Win32 Rectangle
                    IpcUtil.NativeMethods.GetWindowRect(p.MainWindowHandle, rect)

                    Dim rectangle As Rectangle = rect ' Managed Rectangle
                    Return IpcUtil.NativeMethods.MoveWindow(p.MainWindowHandle,
                                                            (rectangle.Left + position.X), (rectangle.Top + position.Y),
                                                            rectangle.Width, rectangle.Height, True)

                End Using

            End If

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Slices the window size of the first ocurrence found of a running process with the specified name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' IpcUtil.UIAutomation.SliceWindowSize("notepad.exe", New Size(+50, -100))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="processName">
        ''' The process name.
        ''' </param>
        ''' 
        ''' <param name="size">
        ''' The new window size.
        ''' </param>
        ''' 
        ''' <param name="throwOnProcessNotFound">
        ''' If <see langword="True"/>, throws an <see cref="ArgumentException"/> exception if any process was found with the specified name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' if successful <see langword="True"/>; if failed <see langword="True"/> or
        ''' if <paramref name="throwOnProcessNotFound"/> is <see langword="False"/> and any process was found with the specified name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Any process found with the specified name.;processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function SliceWindowSize(ByVal processName As String,
                                               ByVal size As Size,
                                               Optional throwOnProcessNotFound As Boolean = False) As Boolean

            Return IpcUtil.UIAutomation.SliceWindowSize(Process.GetProcessesByName(IpcUtil.UIAutomation.FixProcessName(processName)).FirstOrDefault,
                                                        size, throwOnProcessNotFound)

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Slices the window size of the specified process.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' IpcUtil.UIAutomation.SliceWindowSize(Process.GetProcessesByName("notepad").FirstOrDefault, New Size(+100, -50))
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="p">
        ''' The process.
        ''' </param>
        ''' 
        ''' <param name="size">
        ''' The new window size.
        ''' </param>
        ''' 
        ''' <param name="throwOnProcessNotFound">
        ''' If <see langword="True"/>, throws an <see cref="ArgumentException"/> exception if any process was found with the specified name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' if successful <see langword="True"/>; if failed <see langword="True"/> or
        ''' if <paramref name="throwOnProcessNotFound"/> is <see langword="False"/> and any process was found with the specified name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Any process found with the specified name.;processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function SliceWindowSize(ByVal p As Process,
                                               ByVal size As Size,
                                               Optional throwOnProcessNotFound As Boolean = False) As Boolean

            If (p Is Nothing) AndAlso (throwOnProcessNotFound) Then
                Throw New ArgumentException(message:="Any process found with the specified name.", paramName:="processName")

            Else
                Using p

                    Dim rect As IpcUtil.NativeMethods.Rect ' Win32 Rectangle
                    IpcUtil.NativeMethods.GetWindowRect(p.MainWindowHandle, rect)

                    Dim rectangle As Rectangle = rect ' Managed Rectangle
                    Return IpcUtil.NativeMethods.MoveWindow(p.MainWindowHandle,
                                                            rectangle.Left, rectangle.Top,
                                                            rectangle.Width + size.Width, rectangle.Height + size.Height, True)

                End Using

            End If

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Set the visibility state of a window.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim hwnd As IntPtr = Process.GetProcessesByName("notepad").First.MainWindowHandle
        ''' IpcUtil.UIAutomation.SetWindowState(hwnd, IpcUtil.UIAutomation.WindowState.Hide)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwnd">
        ''' A handle to the window.
        ''' </param>
        ''' 
        ''' <param name="windowState">
        ''' The window state.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the window was previously visible, the return value is <see langword="True"/>. 
        ''' If the window was previously hidden, the return value is <see langword="False"/>. 
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function SetWindowState(ByVal hwnd As IntPtr,
                                              ByVal windowState As IpcUtil.UIAutomation.WindowState) As Boolean

            Return IpcUtil.NativeMethods.ShowWindow(hwnd, windowState)

        End Function

        ' ''' ----------------------------------------------------------------------------------------------------
        ' ''' <summary>
        ' ''' Set the visibility state of a window.
        ' ''' </summary>
        ' ''' ----------------------------------------------------------------------------------------------------
        ' ''' <example> This is a code example.
        ' ''' <code>
        ' ''' 
        ' ''' </code>
        ' ''' </example>
        ' ''' ----------------------------------------------------------------------------------------------------
        ' ''' <param name="p">
        ' ''' The process.
        ' ''' </param>
        ' ''' 
        ' ''' <param name="windowState">
        ' ''' The window state.
        ' ''' </param>
        ' ''' ----------------------------------------------------------------------------------------------------
        ' ''' <returns>
        ' ''' If the window was previously visible, the return value is <see langword="True"/>. 
        ' ''' If the window was previously hidden, the return value is <see langword="False"/>. 
        ' ''' </returns>
        ' ''' ----------------------------------------------------------------------------------------------------
        '<DebuggerStepThrough>
        'Public Shared Function SetWindowState(ByVal p As Process,
        '                                      ByVal windowState As IpcUtil.UIAutomation.WindowState) As Boolean

        '    Dim pHandle As IntPtr = IntPtr.Zero
        '    Dim pid As Integer

        '    ' If window is visible then...
        '    If (p.MainWindowHandle <> IntPtr.Zero) Then
        '        Return IpcUtil.NativeMethods.ShowWindow(p.MainWindowHandle, windowState)

        '    Else ' window is hidden.

        '        ' Check all open windows (not only the process we are looking), begining from the child of the desktop.
        '        While (pid <> p.Id)

        '            ' Get child handle of window who's handle is "pHandle".
        '            pHandle = NativeMethods.FindWindowEx(IntPtr.Zero, pHandle, Nothing, Nothing)

        '            ' Get PID from "pHandle".
        '            NativeMethods.GetWindowThreadProcessId(pHandle, pid)

        '        End While

        '        Return NativeMethods.ShowWindow(pHandle, windowState)

        '    End If

        'End Function

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Fixes the name of a process by removing the .exe file extension.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="processName">
        ''' The process name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The fixed process name.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' processName
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Shared Function FixProcessName(ByVal processName As String) As String

            If String.IsNullOrEmpty(processName) Then
                Throw New ArgumentNullException(paramName:="processName")

            Else
                If processName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Then
                    processName = processName.Remove(processName.LastIndexOf(".exe", StringComparison.OrdinalIgnoreCase))
                End If

                Return processName

            End If

        End Function

#End Region

    End Class

#End Region

#End Region

End Module

#End Region
