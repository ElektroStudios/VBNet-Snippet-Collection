
' ***********************************************************************
' Author   : Elektro
' Modified : 28-October-2015
' ***********************************************************************
' <copyright file="Shell Util.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Required References "

' Microsoft Shell Controls And Automation (COM) (Interop.SHDocVw.dll)
' Microsoft Internet Controls             (COM) (Interop.Shell32.dll)

#End Region

#Region " Public Members Summary "

#Region " Child Classes "

' ShellUtil.Desktop
' ShellUtil.Explorer
' ShellUtil.StartMenu
' ShellUtil.TaskBar

#End Region

#Region " Enumerations "

' ShellUtil.TaskBar.TaskbarVisibility As Integer

#End Region

#Region " Properties "

' ShellUtil.Explorer.ExplorerWindows As ReadOnlyCollection(Of ShellBrowserWindow)
' ShellUtil.Explorer.ExplorerWindowsFolders As ReadOnlyCollection(Of Shell32.Folder2)

' ShellUtil.TaskBar.ClassName() As String
' ShellUtil.TaskBar.Hwnd() As Intptr

#End Region

#Region " Methods "

' ShellUtil.Applets.RunDateTime()
' ShellUtil.Applets.RunExecuteDialog()
' ShellUtil.Applets.RunFindComputer()
' ShellUtil.Applets.RunFindFiles()
' ShellUtil.Applets.RunFindPrinter()
' ShellUtil.Applets.RunHelpCenter()
' ShellUtil.Applets.RunSearchCommand()
' ShellUtil.Applets.RunTrayProperties()
' ShellUtil.Applets.RunWindowsSecurity()
' ShellUtil.Applets.RunWindowSwitcher()

' ShellUtil.Desktop.CascadeWindows()
' ShellUtil.Desktop.Hide()
' ShellUtil.Desktop.Show()
' ShellUtil.Desktop.TileWindowsHorizontally()
' ShellUtil.Desktop.TileWindowsVertically()
' ShellUtil.Desktop.ToggleState()

' ShellUtil.Explorer.AddFileToRecentDocs(String)
' ShellUtil.Explorer.RefreshWindows()

' ShellUtil.StartMenu.PinItem(String)
' ShellUtil.StartMenu.UnpinItem(String)

' ShellUtil.TaskBar.Hide(Boolean)
' ShellUtil.TaskBar.PinItem(String)
' ShellUtil.TaskBar.Show(Boolean)
' ShellUtil.TaskBar.UnpinItem(String)

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports SHDocVw
Imports Shell32
Imports System
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices

#End Region

#Region " Shell Util "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains related Windows shell utilities.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module ShellUtil

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
        ''' Retrieves a handle to the top-level window whose class name and window name match the specified strings. 
        ''' This function does not search child windows. 
        ''' This function does not perform a case-sensitive search.
        ''' To search child windows, beginning with a specified child window, use the 'FindWindowEx' function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="lpClassName">
        ''' The class name or a class atom created by a previous call to the RegisterClass or RegisterClassEx function. 
        ''' The atom must be in the low-order word of <paramref name="lpClassName"/>; the high-order word must be zero.
        ''' If <paramref name="lpClassName"/> is NULL, it finds any window whose title matches the <paramref name="lpWindowName"/> parameter.
        ''' </param>
        ''' 
        ''' <param name="lpWindowName">
        ''' The window name (the window's titlebar title). 
        ''' If this parameter is <see langword="Nothing"/>, all window names match.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a handle to the window that has the specified class name and window name.
        ''' If the function fails, the return value is <see langword="Nothing"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms633499%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function FindWindow(
                               ByVal lpClassName As String,
                               ByVal lpWindowName As String
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the specified window's show state.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hwnd">
        ''' A handle to the window.
        ''' </param>
        ''' 
        ''' <param name="nCmdShow">
        ''' Controls how the window is to be shown.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the window was previously visible, the return value is nonzero.
        ''' If the window was previously hidden, the return value is zero.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll")>
        Friend Shared Function ShowWindow(
                               ByVal hwnd As IntPtr,
                               ByVal nCmdShow As Integer
        ) As Integer
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
        Friend Shared Function SendMessage(
                               ByVal hWnd As IntPtr,
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
            ''' Message sent when the user selects a command item from a menu,
            ''' when a control sends a notification message to its parent window,
            ''' or when an accelerator keystroke is translated.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
            ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms647591%28v=vs.85%29.aspx"/>
            ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            WMCommand = &H111UI

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

            ''' <summary>
            ''' Minimize all windows.
            ''' Used with <see cref="ShellUtil.NativeMethods.WindowsMessages.WmCommand"/> message.
            ''' </summary>
            MinimizeAll = 419UI

            ''' <summary>
            ''' Undo the minimization of all minimized windows.
            ''' Used with <see cref="ShellUtil.NativeMethods.WindowsMessages.WmCommand"/> message.
            ''' </summary>
            UndoMinimizeAll = 416UI

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

#Region " Applets "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related Windows applet utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class Applets

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="ShellUtil.Applets"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Taskbar and Start Menu Properties dialog box. 
        ''' This method has the same effect as right-clicking the taskbar and selecting Properties.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774105%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunTrayProperties()

            Dim shell As New Shell32.Shell
            shell.TrayProperties()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Date and Time Properties dialog box. 
        ''' This method has the same effect as right-clicking the clock in the taskbar status area and selecting Adjust Date/Time.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774092%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunDateTime()

            Dim shell As New Shell32.Shell
            shell.SetTime()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Apps Search pane.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/jj635751%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunSearchCommand()

            Dim shell As New Shell32.Shell
            shell.SearchCommand()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Invokes the Window Switcher (ALT+TAB).
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/gg537749%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunWindowSwitcher()

            Dim shell As New Shell32.Shell
            shell.WindowSwitcher()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Run dialog to the user. 
        ''' This method has the same effect as clicking the Start menu and selecting Run.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774075%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunExecuteDialog()

            Dim shell As New Shell32.Shell
            shell.FileRun()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Search Results: Computers dialog box. 
        ''' The dialog box shows the result of the search for a specified computer.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774077%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunFindComputer()

            Dim shell As New Shell32.Shell
            shell.FindComputer()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Find: All Files dialog box. 
        ''' This is the same as clicking the Start menu and then selecting "Search".
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774079%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunFindFiles()

            Dim shell As New Shell32.Shell
            shell.FindFiles()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Find Printer dialog box.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/gg537738%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunFindPrinter()

            Dim shell As New Shell32.Shell
            shell.FindPrinter()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Windows Help and Support Center. 
        ''' This method has the same effect as clicking the Start menu and selecting Help and Support.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774081%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunHelpCenter()

            Dim shell As New Shell32.Shell
            shell.Help()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Displays the Windows Security dialog box.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/gg537748%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RunWindowsSecurity()

            Dim shell As New Shell32.Shell
            shell.WindowsSecurity()
            Marshal.ReleaseComObject(shell)

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

#Region " Desktop "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related Windows desktop utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class Desktop

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="ShellUtil.Desktop"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Shows/Restores the desktop.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub Show()

            ShellUtil.NativeMethods.SendMessage(
                ShellUtil.NativeMethods.FindWindow(ShellUtil.TaskBar.ClassName, String.Empty),
                ShellUtil.NativeMethods.WindowsMessages.WMCommand,
                New IntPtr(ShellUtil.NativeMethods.WParams.UndoMinimizeAll),
                New IntPtr(ShellUtil.NativeMethods.LParams.None))

            '' Old methodology:
            ' Dim shell As New Shell
            ' shell.UndoMinimizeALL()
            ' Marshal.ReleaseComObject(shell)


        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Hides the desktop.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub Hide()

            ShellUtil.NativeMethods.SendMessage(
                ShellUtil.NativeMethods.FindWindow(ShellUtil.TaskBar.ClassName, String.Empty),
                ShellUtil.NativeMethods.WindowsMessages.WMCommand,
                New IntPtr(ShellUtil.NativeMethods.WParams.MinimizeAll),
                New IntPtr(ShellUtil.NativeMethods.LParams.None))

            '' Old methodology:
            ' Dim shell As New Shell
            ' shell.MinimizeAll()
            ' Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Toggles the visibility state of the desktop.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="NotImplementedException">
        ''' This feature is not supported under Windows XP.
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub ToggleState()

            ' If current O.S = Windows Vista or above, then...
            If Environment.OSVersion.Version.Major >= 6 Then
                Dim shell As New Shell32.Shell
                shell.ToggleDesktop()
                Marshal.ReleaseComObject(shell)

            Else
                Throw New NotImplementedException(message:="This feature is not supported under Windows XP.")

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Cascades all of the windows on the desktop. 
        ''' This method has the same effect as right-clicking the taskbar and selecting Cascade Windows.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774067%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub CascadeWindows()

            Dim shell As New Shell32.Shell
            shell.CascadeWindows()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Tiles all of the windows on the desktop horizontally. 
        ''' This method has the same effect as right-clicking the taskbar and selecting Tile Windows Horizontally.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774102%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub TileWindowsHorizontally()

            Dim shell As New Shell32.Shell
            shell.TileHorizontally()
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Tiles all of the windows on the desktop vertically. 
        ''' This method has the same effect as right-clicking the taskbar and selecting Tile Windows Vertically.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb774104%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub TileWindowsVertically()

            Dim shell As New Shell32.Shell
            shell.TileVertically()
            Marshal.ReleaseComObject(shell)

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

#Region " Explorer "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related Windows Explorer utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class Explorer

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="ReadOnlyCollection(Of ShellBrowserWindow)"/> containing the opened windows explorer instances.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A <see cref="ReadOnlyCollection(Of ShellBrowserWindow)"/> containing the opened windows explorer instances.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property ExplorerWindows() As ReadOnlyCollection(Of ShellBrowserWindow)
            <DebuggerStepThrough>
            Get
                Return New ReadOnlyCollection(Of ShellBrowserWindow)(ShellUtil.Explorer.GetExplorerWindows.ToList)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="ReadOnlyCollection(Of folder2)"/> containing the opened windows explorer folder instances.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A <see cref="ReadOnlyCollection(Of folder2)"/> containing the opened windows explorer folder instances.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property ExplorerWindowsFolders() As ReadOnlyCollection(Of Shell32.Folder2)
            <DebuggerStepThrough>
            Get
                Return New ReadOnlyCollection(Of Folder2)(ShellUtil.Explorer.GetExplorerWindowsFolders.ToList)
            End Get
        End Property

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="ShellUtil.Explorer"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Adds an item into the Windows recent docs list (MRU).
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filePath">
        ''' The file path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' filePath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/gg537735%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub AddFileToRecentDocs(ByVal filePath As String)

            If String.IsNullOrWhiteSpace(filePath) Then
                Throw New ArgumentNullException(paramName:="filePath")

            Else
                Dim shell As New Shell32.Shell
                shell.AddToRecent(filePath)
                Marshal.ReleaseComObject(shell)

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Refreshes the opened windows explorer folder instances.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RefreshWindows()

            For Each window As ShellBrowserWindow In ShellUtil.Explorer.GetExplorerWindows()
                window.Refresh()
            Next window

        End Sub

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="IEnumerable(Of ShellBrowserWindow)"/> containing the opened windows explorer instances.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' For Each window As ShellBrowserWindow In GetExplorerWindows()
        '''     Console.WriteLine(window.LocationURL)
        ''' Next window
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="IEnumerable(Of ShellBrowserWindow)"/> containing the opened windows explorer instances.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="NotImplementedException">
        ''' Detected an unknown window type.
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Shared Iterator Function GetExplorerWindows() As IEnumerable(Of ShellBrowserWindow)

            Dim shell As New Shell32.Shell

            Try
                For Each window As ShellBrowserWindow In DirectCast(shell.Windows, IShellWindows)

                    Select Case window.Document.GetType.Name

                        Case "HTMLDocumentClass" ' Internet Explorer Window.
                            ' Do Nothing.

                        Case "__ComObject" ' Explorer Window.
                            Yield window

                        Case Else ' Unknown window.
                            Throw New NotImplementedException("Detected an unknown window type.")

                    End Select

                Next window

            Catch ex As Exception
                Throw ex

            Finally
                Marshal.ReleaseComObject(shell)

            End Try

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="IEnumerable(Of Folder2)"/> containing the opened windows explorer folder instances.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' For Each folder As Folder2 In GetExplorerWindowsFolders()
        '''     Console.WriteLine(folder.Self.Path)
        ''' Next folder
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="IEnumerable(Of Folder2)"/> containing the opened windows explorer folder instances.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="NotImplementedException">
        ''' Detected an unknown window type.
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Shared Iterator Function GetExplorerWindowsFolders() As IEnumerable(Of Folder2)

            Dim shell As New Shell32.Shell

            Try
                For Each window As ShellBrowserWindow In DirectCast(shell.Windows, IShellWindows)

                    Select Case window.Document.GetType.Name

                        Case "HTMLDocumentClass" ' Internet Explorer Window.
                            ' Do Nothing.

                        Case "__ComObject" ' Explorer Window.
                            Yield DirectCast(DirectCast(window.Document, ShellFolderView).Folder, Folder2)

                        Case Else ' Unknown window.
                            Throw New NotImplementedException("Detected an unknown window type.")

                    End Select

                Next window

            Catch ex As Exception
                Throw ex

            Finally
                Marshal.ReleaseComObject(shell)

            End Try

        End Function

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

#Region " StartMenu "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related Windows startmenu utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class StartMenu

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="ShellUtil.StartMenu"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Pins an item on the startmenu.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemPath">
        ''' The file or directory path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub PinItem(ByVal itemPath As String)

            Dim shell As New Shell32.Shell
            Dim link As FolderItem = shell.NameSpace(Path.GetDirectoryName(itemPath)).ParseName(Path.GetFileName(itemPath))

            ' HKEY_CURRENT_USER\Software\Classes\CLSID\{a2a9545d-a0c2-42b4-9708-a0b2badd77c9}
            link.InvokeVerb("startpin")
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Unpins an item from the startmenu.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemPath">
        ''' The file or directory path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub UnpinItem(ByVal itemPath As String)

            Dim shell As New Shell32.Shell
            Dim link As FolderItem = shell.NameSpace(Path.GetDirectoryName(itemPath)).ParseName(Path.GetFileName(itemPath))

            ' HKEY_CURRENT_USER\Software\Classes\CLSID\{a2a9545d-a0c2-42b4-9708-a0b2badd77c9}
            link.InvokeVerb("startunpin")
            Marshal.ReleaseComObject(shell)

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

#Region " TaskBar "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related Windows desktop's taskbar utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class TaskBar

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the taskbar class name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The taskbar class name.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property ClassName As String
            Get
                Return "Shell_TrayWnd"
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the taskbar window handle.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The taskbar window handle
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property Hwnd As IntPtr
            <DebuggerStepThrough>
            Get
                Return ShellUtil.NativeMethods.FindWindow(ShellUtil.TaskBar.ClassName, Nothing)
            End Get
        End Property

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Specifies a desktop taskbar visibility flag.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Enum TaskBarVisibility As Integer

            ''' <summary>
            ''' Hides the TaskBar.
            ''' </summary>
            Hide = &H0I

            ''' <summary>
            ''' Shows the TaskBar.
            ''' </summary>
            Show = &H5I

        End Enum

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="ShellUtil.TaskBar"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Hides the desktop taskbar.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Exception">
        ''' The taskbar was already hidden.
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub Hide(Optional ByVal ignoreErrors As Boolean = True)

            If (ShellUtil.TaskBar.SetVisibility(ShellUtil.TaskBar.TaskBarVisibility.Hide) = 0) AndAlso
               (Not ignoreErrors) Then

                Throw New Exception("The taskbar was already hidden.")

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Shows the desktop taskbar.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Exception">
        ''' The taskbar was already shown.
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub Show(Optional ByVal ignoreErrors As Boolean = True)

            If (ShellUtil.TaskBar.SetVisibility(ShellUtil.TaskBar.TaskBarVisibility.Show) <> 0) AndAlso
               (Not ignoreErrors) Then

                Throw New Exception("The taskbar was already shown.")

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Pins an item on TaskBar.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemPath">
        ''' The file or directory path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub PinItem(ByVal itemPath As String)

            Dim shell As New Shell32.Shell
            Dim link As FolderItem = shell.NameSpace(Path.GetDirectoryName(itemPath)).ParseName(Path.GetFileName(itemPath))

            ' HKEY_CLASSES_ROOT\CLSID\{90AA3A4E-1CBA-4233-B8BB-535773D48449}
            link.InvokeVerb("taskbarpin")
            Marshal.ReleaseComObject(shell)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Unpins an item from TaskBar.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemPath">
        ''' The file or directory path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub UnpinItem(ByVal itemPath As String)

            Dim shell As New Shell32.Shell
            Dim link As FolderItem = shell.NameSpace(Path.GetDirectoryName(itemPath)).ParseName(Path.GetFileName(itemPath))

            ' HKEY_CLASSES_ROOT\CLSID\{90AA3A4E-1CBA-4233-B8BB-535773D48449}
            link.InvokeVerb("taskbarunpin")
            Marshal.ReleaseComObject(shell)

        End Sub

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the Windows TaskBar visibility.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="visibility">
        ''' The desired TaskBar visibility.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the window was previously visible, the return value is nonzero.
        ''' If the window was previously hidden, the return value is zero.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Shared Function SetVisibility(ByVal visibility As ShellUtil.TaskBar.TaskBarVisibility) As Integer

            Return ShellUtil.NativeMethods.ShowWindow(ShellUtil.TaskBar.Hwnd, visibility)

        End Function

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
