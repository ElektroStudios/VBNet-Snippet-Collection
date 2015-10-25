' ***********************************************************************
' Author   : Elektro
' Modified : 25-October-2015
' ***********************************************************************
' <copyright file="EnvironmentUtil.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " ToDo: "

' Implement more API wrappers to supply more member clones of the rest of "System.Windows.Forms.SystemInformation" properties, with setters.

#End Region

#Region " Required References "

' Microsoft Shell Controls And Automation (COM) (Interop.SHDocVw.dll)
' Microsoft Internet Controls             (COM) (Interop.Shell32.dll)

#End Region

#Region " Public Members Summary "

#Region " Child Classes "

' EnvironmentUtil.EnvironmentVariables
' EnvironmentUtil.FileSystem
' EnvironmentUtil.OS
' EnvironmentUtil.Programs
' EnvironmentUtil.Shell
' EnvironmentUtil.Shell.Desktop
' EnvironmentUtil.Shell.Explorer
' EnvironmentUtil.Shell.StartMenu
' EnvironmentUtil.Shell.TaskBar
' EnvironmentUtil.Theming

#End Region

#Region " Enumerations "

' EnvironmentUtil.EnvironmentScope
' EnvironmentUtil.OS.Architecture
' EnvironmentUtil.Theming.CursorType
' EnvironmentUtil.Theming.WallpaperStyle

#End Region

#Region " Properties "

' EnvironmentUtil.EnvironmentVariables.CurrentVariables(EnvironmentUtil.EnvironmentScope) As ReadOnlyCollection(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)
' EnvironmentUtil.OS.ActiveWindowTrackingEnabled As Boolean
' EnvironmentUtil.OS.ActiveWindowTrackingTimeout As UShort
' EnvironmentUtil.OS.BeepEnabled As Boolean
' EnvironmentUtil.OS.BlockSendInputResetsEnabled As Boolean
' EnvironmentUtil.OS.BorderMultiplierFactor As Integer
' EnvironmentUtil.OS.CaretWidth As Integer
' EnvironmentUtil.OS.CleartypeEnabled As Boolean
' EnvironmentUtil.OS.ClientAreaAnimationEnabled As Boolean
' EnvironmentUtil.OS.ComboBoxAnimationEnabled As Boolean
' EnvironmentUtil.OS.CurrentArchitecture() As EnvironmentUtil.OS.Architecture
' EnvironmentUtil.OS.CursorShadowEnabled As Boolean
' EnvironmentUtil.OS.DoubleClickSize As Size
' EnvironmentUtil.OS.DoubleClickTime As Integer
' EnvironmentUtil.OS.DragFullWindowsEnabled As Boolean
' EnvironmentUtil.OS.DragSize As Size
' EnvironmentUtil.OS.DropShadowEnabled As Boolean
' EnvironmentUtil.OS.FlatMenuEnabled As Boolean
' EnvironmentUtil.OS.FocusBorderSize As Size
' EnvironmentUtil.OS.FontSmoothingContrast As Integer
' EnvironmentUtil.OS.FontSmoothingEnabled As Boolean
' EnvironmentUtil.OS.ForegroundFlashCount As UShort
' EnvironmentUtil.OS.ForegroundLockTimeout As UShort
' EnvironmentUtil.OS.HotTrackingEnabled As Boolean
' EnvironmentUtil.OS.HungAppTimeout As Integer
' EnvironmentUtil.OS.IconSpacing As Size
' EnvironmentUtil.OS.IconTitleWrappingEnabled As Boolean
' EnvironmentUtil.OS.KeyboardDelay As Integer
' EnvironmentUtil.OS.KeyboardSpeed As Integer
' EnvironmentUtil.OS.ListBoxSmoothScrollingEnabled As Boolean
' EnvironmentUtil.OS.MenuAccessKeysUnderlined As Boolean
' EnvironmentUtil.OS.MenuAnimationEnabled As Boolean
' EnvironmentUtil.OS.MenuFadeEnabled As Boolean
' EnvironmentUtil.OS.MenuShowDelay As Integer
' EnvironmentUtil.OS.MessageDuration As Long
' EnvironmentUtil.OS.MouseButtonsSwapEnabled As Boolean
' EnvironmentUtil.OS.MouseClickLockEnabled As Boolean
' EnvironmentUtil.OS.MouseClickLockTime As Integer
' EnvironmentUtil.OS.MouseHoverSize As Size
' EnvironmentUtil.OS.MouseHoverTime As Integer
' EnvironmentUtil.OS.MouseSonarEnabled As Boolean
' EnvironmentUtil.OS.MouseSpeed As Integer
' EnvironmentUtil.OS.MouseTrailAmount As Integer
' EnvironmentUtil.OS.MouseVanishEnabled As Boolean
' EnvironmentUtil.OS.MouseWheelScrollLines As Integer
' EnvironmentUtil.OS.OverlappedContentEnabled As Boolean
' EnvironmentUtil.OS.PopupMenuAlignment As LeftRightAlignment
' EnvironmentUtil.OS.ScreensaverEnabled As Boolean
' EnvironmentUtil.OS.ScreensaverPath As String
' EnvironmentUtil.OS.ScreensaverTimeout As Integer
' EnvironmentUtil.OS.ScreensaveSecureEnabled As Boolean
' EnvironmentUtil.OS.SelectionFadeEnabled As Boolean
' EnvironmentUtil.OS.SnapToDefaultEnabled As Boolean
' EnvironmentUtil.OS.SystemDateTime As Date
' EnvironmentUtil.OS.SystemLanguageBarEnabled As Boolean
' EnvironmentUtil.OS.TitleBarGradientEnabled As Boolean
' EnvironmentUtil.OS.ToolTipAnimationEnabled As Boolean
' EnvironmentUtil.OS.UIEffectsEnabled As Boolean
' EnvironmentUtil.OS.WaitToKillAppTimeout As Integer
' EnvironmentUtil.OS.WaitToKillServiceTimeout As Integer
' EnvironmentUtil.OS.WheelscrollChars As Integer
' EnvironmentUtil.Programs.DefaultWebBrowser() As String
' EnvironmentUtil.Programs.IExplorerVersion() As Version
' EnvironmentUtil.Shell.Explorer.ExplorerWindows As ReadOnlyCollection(Of ShellBrowserWindow)
' EnvironmentUtil.Shell.Explorer.ExplorerWindowsFolders As ReadOnlyCollection(Of Shell32.Folder2)
' EnvironmentUtil.Shell.TaskBar.ClassName() As String
' EnvironmentUtil.Shell.TaskBar.Hwnd() As Intptr
' EnvironmentUtil.Theming.AeroEnabled() As Boolean
' EnvironmentUtil.Theming.AeroSupported() As Boolean
' EnvironmentUtil.Theming.CurrentTheme() As EnvironmentUtil.Theming.ThemeInfo
' EnvironmentUtil.Theming.CurrentWallpaper() As String
' EnvironmentUtil.Theming.WallpaperAsJpegIsSupported() As Boolean
' EnvironmentUtil.Theming.WallpaperStylesFitFillAreSupported() As Boolean

#End Region

#Region " Types "

' EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo
' EnvironmentUtil.Theming.ThemeInfo

#End Region

#Region " Functions "

' EnvironmentUtil.EnvironmentVariables.GetValue(EnvironmentUtil.EnvironmentScope, String, Boolean) As String
' EnvironmentUtil.EnvironmentVariables.GetVariableInfo(EnvironmentUtil.EnvironmentScope, String, Boolean) As EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo
' EnvironmentUtil.FileSystem.GetItemVerbs(String) As IEnumerable(Of FolderItemVerb)
' EnvironmentUtil.FileSystem.ItemNameIsInvalid(String) As Boolean
' EnvironmentUtil.FileSystem.ItemNameOrPathIsInvalid(String) As Boolean
' EnvironmentUtil.FileSystem.ItemPathIsInvalid(String) As Boolean

#End Region

#Region " Methods "

' EnvironmentUtil.EnvironmentVariables.RegisterVariable(EnvironmentUtil.EnvironmentScope, EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo, Boolean)
' EnvironmentUtil.EnvironmentVariables.RegisterVariable(EnvironmentUtil.EnvironmentScope, String, String, Boolean)
' EnvironmentUtil.EnvironmentVariables.UnregisterVariable(EnvironmentUtil.EnvironmentScope, String, Boolean)
' EnvironmentUtil.FileSystem.InvokeItemVerb(String, String)
' EnvironmentUtil.OS.NotifyDirectoryAttributesChanged(String)
' EnvironmentUtil.OS.NotifyDirectoryCreated(String)
' EnvironmentUtil.OS.NotifyDirectoryDeleted(String)
' EnvironmentUtil.OS.NotifyDirectoryRenamed(String, String)
' EnvironmentUtil.OS.NotifyDriveAdded(String, Boolean)
' EnvironmentUtil.OS.NotifyDriveRemoved(String)
' EnvironmentUtil.OS.NotifyFileAssociationChanged()
' EnvironmentUtil.OS.NotifyFileAttributesChanged(String)
' EnvironmentUtil.OS.NotifyFileCreated(String)
' EnvironmentUtil.OS.NotifyFileDeleted(String)
' EnvironmentUtil.OS.NotifyFileRenamed(String, String)
' EnvironmentUtil.OS.NotifyFreespaceChanged(String)
' EnvironmentUtil.OS.NotifyMediaInserted(String)
' EnvironmentUtil.OS.NotifyMediaRemoved(String)
' EnvironmentUtil.OS.NotifyNetworkFolderShared(String)
' EnvironmentUtil.OS.NotifyNetworkFolderUnshared(String)
' EnvironmentUtil.OS.NotifyUpdateDirectory(String)
' EnvironmentUtil.OS.NotifyUpdateImage()
' EnvironmentUtil.OS.ReloadSystemCursors()
' EnvironmentUtil.OS.ReloadSystemIcons()
' EnvironmentUtil.OS.RunDateTime()
' EnvironmentUtil.OS.RunExecuteDialog()
' EnvironmentUtil.OS.RunFindComputer()
' EnvironmentUtil.OS.RunFindFiles()
' EnvironmentUtil.OS.RunFindPrinter()
' EnvironmentUtil.OS.RunHelpCenter()
' EnvironmentUtil.OS.RunSearchCommand()
' EnvironmentUtil.OS.RunTrayProperties()
' EnvironmentUtil.OS.RunWindowsSecurity()
' EnvironmentUtil.OS.RunWindowSwitcher()
' EnvironmentUtil.Shell.Desktop.CascadeWindows()
' EnvironmentUtil.Shell.Desktop.Hide()
' EnvironmentUtil.Shell.Desktop.Show()
' EnvironmentUtil.Shell.Desktop.TileWindowsHorizontally()
' EnvironmentUtil.Shell.Desktop.TileWindowsVertically()
' EnvironmentUtil.Shell.Desktop.ToggleState()
' EnvironmentUtil.Shell.Explorer.AddFileToRecentDocs(String)
' EnvironmentUtil.Shell.Explorer.RefreshWindows()
' EnvironmentUtil.Shell.StartMenu.PinItem(String)
' EnvironmentUtil.Shell.StartMenu.UnpinItem(String)
' EnvironmentUtil.Shell.TaskBar.Hide(Boolean)
' EnvironmentUtil.Shell.TaskBar.PinItem(String)
' EnvironmentUtil.Shell.TaskBar.Show(Boolean)
' EnvironmentUtil.Shell.TaskBar.UnpinItem(String)
' EnvironmentUtil.Theming.RemoveDesktopWallpaper()
' EnvironmentUtil.Theming.SetDesktopWallpaper(String, EnvironmentUtil.Theming.WallpaperStyle)
' EnvironmentUtil.Theming.SetSystemCursor(String, EnvironmentUtil.Theming.CursorType)
' EnvironmentUtil.Theming.SetSystemVisualTheme(String, String, String)

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports Microsoft.Win32
Imports SHDocVw
Imports Shell32
Imports System
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.Text
Imports System.Windows.Forms

#End Region

#Region " Environment Util "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains related Windows environment utilities.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public NotInheritable Class EnvironmentUtil

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
        Friend Shared Function SendMessage(
                               ByVal hWnd As IntPtr,
                               ByVal msg As WindowsMessages,
                               ByVal wParam As IntPtr,
                               ByVal lParam As IntPtr
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Enables an application to customize the system cursors. 
        ''' It replaces the contents of the system cursor specified by the <paramfer name="id"/> parameter 
        ''' with the contents of the cursor specified by the <paramfer name="hCursor"/> parameter and then destroys <paramfer name="hCursor"/>. 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hCursor">
        ''' A handle to the cursor. 
        ''' The function replaces the contents of the system cursor specified by <paramfer name="id"/> parameter  
        ''' with the contents of the cursor handled by <paramfer name="hCursor"/> parameter.
        ''' </param>
        ''' 
        ''' <param name="id">
        ''' The system cursor to replace with the contents of <paramfer name="hCursor"/> parameter.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is <see langword="True"/>.
        ''' If the function fails, the return value is <see langword="False"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms648395%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True)>
        Friend Shared Function SetSystemCursor(
                               ByVal hCursor As IntPtr,
                               ByVal id As UInteger
        ) As Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates a cursor based on data contained in a file. 
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="lpFileName">
        ''' The source of the file data to be used to create the cursor. 
        ''' The data in the file must be in either .CUR or .ANI format.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is an <see cref="IntPtr"/> handle to the new cursor.
        ''' If the function fails, the return value is <see cref="IntPtr.Zero"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms648392%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True, charSet:=CharSet.Ansi, bestFitMapping:=False, throwOnUnmappableChar:=True)>
        Friend Shared Function LoadCursorFromFile(
                               ByVal lpFileName As String
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sends the specified message to one or more windows.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' A handle to the window whose window procedure will receive the message.
        ''' 
        ''' If this parameter is <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.HWND_BROADCAST"/>, 
        ''' the message is sent to all top-level windows in the system, including disabled or invisible unowned windows.
        ''' 
        ''' The function does not return until each window has timed out. 
        ''' Therefore, the total wait time can be up to the value of uTimeout multiplied by the number of top-level windows.
        ''' </param>
        ''' 
        ''' <param name="Msg">
        ''' The message to be sent.
        ''' For lists of the system-provided messages, see System-Defined Messages:
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644927%28v=vs.85%29.aspx#system_defined"/>
        ''' </param>
        ''' 
        ''' <param name="wParam">
        ''' Any additional message-specific information.
        ''' </param>
        ''' 
        ''' <param name="lParam">
        ''' Any additional message-specific information.
        ''' </param>
        ''' 
        ''' <param name="fuFlags">
        ''' The behavior of this function.
        ''' </param>
        ''' 
        ''' <param name="uTimeout">
        ''' The duration of the time-out period, in milliseconds. 
        ''' If the message is a broadcast message, each window can use the full time-out period. 
        ''' For example, if you specify a five second time-out period and there are three top-level windows that fail to process the message, 
        ''' you could have up to a 15 second delay.
        ''' </param>
        ''' 
        ''' <param name="lpdwResult">
        ''' The result of the message processing. 
        ''' The value of this parameter depends on the message that is specified.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is nonzero. 
        ''' 
        ''' If the function fails or times out, the return value is 0.
        ''' 
        ''' <see cref="EnvironmentUtil.NativeMethods.SendMessageTimeout"/> does not provide information about 
        ''' individual windows timing out if <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.HWND_BROADCAST"/> is used.
        ''' 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' If <see cref="Marshal.GetLastWin32Error"/> returns ERROR_TIMEOUT, then the function timed out.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644952%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SendMessageTimeout(
                               ByVal hwnd As IntPtr,
                               ByVal msg As Integer,
                               ByVal wParam As IntPtr,
                               ByVal lParam As String,
                               ByVal fuFlags As SendMessageTimeoutFlags,
                               ByVal uTimeout As Integer,
                         <Out> ByRef lpdwResult As IntPtr
        ) As IntPtr
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the name of the current visual style, and optionally retrieves the color scheme name and size name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pszThemeFileName">
        ''' Pointer to a string that receives the theme path and file name.
        ''' </param>
        ''' 
        ''' <param name="dwMaxNameChars">
        ''' The maximum number of characters allowed in the theme file name.
        ''' </param>
        ''' 
        ''' <param name="pszColorBuff">
        ''' Pointer to a string that receives the color scheme name. This parameter may be set to <see langword="Nothing"/>.
        ''' </param>
        ''' 
        ''' <param name="cchMaxColorChars">
        ''' The maximum number of characters allowed in the color scheme name.
        ''' </param>
        ''' 
        ''' <param name="pszSizeBuff">
        ''' Pointer to a string that receives the size name. This parameter may be set to NULL.</param>
        ''' 
        ''' <param name="cchMaxSizeChars">
        ''' The maximum number of characters allowed in the size name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns '0' if successful, otherwise, an error code.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb773365%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("uxtheme", CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function GetCurrentThemeName(
                               ByVal pszThemeFileName As StringBuilder,
                               ByVal dwMaxNameChars As Integer,
                               ByVal pszColorBuff As StringBuilder,
                               ByVal cchMaxColorChars As Integer,
                               ByVal pszSizeBuff As StringBuilder,
                               ByVal cchMaxSizeChars As Integer
        ) As Integer
        End Function

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
        ''' Obtains a value that indicates whether Desktop Window Manager (DWM) composition is enabled. 
        ''' Applications on machines running Windows 7 or earlier can listen for composition state changes by handling the WM_DWMCOMPOSITIONCHANGED notification.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="enabled">
        ''' A pointer to a value that, when this function returns successfully, receives <see langword="True"/> if DWM composition is enabled; otherwise, <see langword="False"/>.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' this function succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/aa969518%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("dwmapi.dll")>
        Friend Shared Function DwmIsCompositionEnabled(
                               ByRef enabled As Boolean
        ) As Integer
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As UInteger,
                               ByVal pvParam As String,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As UInteger,
                  <[In]> <Out> ByVal pvParam As StringBuilder,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As UInteger,
                               ByRef pvParam As Boolean,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As Boolean,
                               ByVal pvParam As UInteger,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As UInteger,
                               ByRef pvParam As UInteger,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As Integer,
                               ByVal pvParam As UInteger,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As UInteger,
                               ByVal pvParam As Integer,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As Integer,
                               ByVal pvParam As Boolean,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As UInteger,
                               ByRef pvParam As UShort,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As UInteger,
                               ByRef pvParam As ULong,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves or sets the value of one of the system-wide parameters.
        ''' This function can also update the user profile while setting a parameter.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="uiAction">
        ''' The system-wide parameter to be retrieved or set.
        ''' </param>
        ''' 
        ''' <param name="uiParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify '0' for this parameter.
        ''' </param>
        ''' 
        ''' <param name="pvParam">
        ''' A parameter whose usage and format depends on the system parameter being queried or set. 
        ''' For more information about system-wide parameters, see the <paramfer name="uiAction"></paramfer> parameter. 
        ''' If not otherwise indicated, you must specify <see langword="Nothing"/> for this parameter.
        ''' For information on the PVOID datatype, see Windows Data Types.
        ''' </param>
        ''' 
        ''' <param name="fWinIni">
        ''' If a system parameter is being set, specifies whether the user profile is to be updated, 
        ''' and if so, whether the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' If the function succeeds, the return value is a nonzero value.
        ''' If the function fails, the return value is zero. 
        ''' To get extended error information, call <see cref="Marshal.GetLastWin32Error"/>.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("user32.dll", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Function SystemParametersInfo(
                               ByVal uiAction As SystemParametersActionFlags,
                               ByVal uiParam As UInteger,
                               ByVal pvParam As Long,
                               ByVal fWinIni As SystemParametersWinIniFlags
        ) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the system visual theme.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="pszThemeFileName">
        ''' The theme filepath (themme.msstyles).
        ''' </param>
        ''' 
        ''' <param name="pszColor">
        ''' The color scheme name.
        ''' </param>
        ''' 
        ''' <param name="pszSize">
        ''' The size name.</param>
        ''' 
        ''' <param name="dwReserved">
        ''' Reserved parameter by the system.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns></returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://www.pinvoke.net/default.aspx/uxtheme/SetSystemVisualStyle.html"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("UxTheme.DLL", BestFitMapping:=False, CallingConvention:=CallingConvention.Winapi, CharSet:=CharSet.Unicode, EntryPoint:="#65")>
        Friend Shared Function SetSystemVisualStyle(
                               ByVal pszThemeFileName As String,
                               ByVal pszColor As String,
                               ByVal pszSize As String,
                               ByVal dwReserved As Integer
        ) As Integer
        End Function

#End Region

#Region " Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system of an event that an application has performed. 
        ''' An application should use this function if it performs an action that may affect the Shell.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="wEventId">
        ''' Describes the event that has occurred. 
        ''' Typically, only one event is specified at a time. 
        ''' If more than one event is specified, the values contained in the dwItem1 and dwItem2 parameters must be the same, respectively, for all specified events.
        ''' </param>
        ''' 
        ''' <param name="uFlags">
        ''' Flags that, when combined bitwise with SHCNF_TYPE, 
        ''' indicate the meaning of the dwItem1 and dwItem2 parameters.
        ''' </param>
        ''' 
        ''' <param name="dwItem1">
        ''' Optional. 
        ''' First event-dependent value.
        ''' </param>
        ''' 
        ''' <param name="dwItem2">
        ''' Optional.
        ''' Second event-dependent value.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762118%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("shell32.dll", EntryPoint:="SHChangeNotify", SetLastError:=True, CharSet:=CharSet.Auto, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
        Friend Shared Sub SHChangeNotify(
                      ByVal wEventId As EnvironmentUtil.NativeMethods.SHChangeNotifyEventID,
                      ByVal uFlags As EnvironmentUtil.NativeMethods.SHChangeNotifyFlags,
                      ByVal dwItem1 As String,
                      ByVal dwItem2 As String)
        End Sub

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="fuFlags"/> parameter of <see cref="EnvironmentUtil.NativeMethods.SendMessageTimeout"/> function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644952%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <Flags()>
        Friend Enum SendMessageTimeoutFlags As Integer

            ''' <summary>
            ''' The calling thread is not prevented from processing other requests while waiting for the function to return.
            ''' </summary>
            Normal = &H0

            ''' <summary>
            ''' Prevents the calling thread from processing any other requests until the function returns.
            ''' </summary>
            Block = &H1

            ''' <summary>
            ''' The function returns without waiting for the time-out period to elapse if the receiving thread appears to not respond or "hangs."
            ''' </summary>
            AbortIfHung = &H2

            ''' <summary>
            ''' The function does not enforce the time-out period  as long as the receiving thread is processing messages.
            ''' </summary>
            NoTimeoutIfNotHung = &H8

            ''' <summary>
            ''' The function should return 0 if the receiving window is destroyed or its owning thread dies while the message is being processed.
            ''' </summary>
            ErrorOnExit = &H20

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="fWinIni"/> parameter of <see cref="EnvironmentUtil.NativeMethods.SetSystemCursor"/> function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms648395%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Enum SystemCursorId As UInteger

            ''' <summary>
            ''' Standard arrow and small hourglass.
            ''' </summary>
            AppStarting = 32650UI

            ''' <summary>
            ''' Standard arrow.
            ''' </summary>
            Arrow = 32512UI

            ''' <summary>
            ''' Crosshair.
            ''' </summary>
            Crosshair = 32515UI

            ''' <summary>
            ''' Hand.
            ''' </summary>
            Hand = 32649UI

            ''' <summary>
            ''' Arrow and question mark.
            ''' </summary>
            Help = 32651UI

            ''' <summary>
            ''' I-beam.
            ''' </summary>
            IBeam = 32513UI

            ''' <summary>
            ''' Slashed circle.
            ''' </summary>
            No = 32648UI

            ''' <summary>
            ''' Four-pointed arrow pointing north, south, east, and west.
            ''' </summary>
            SizeAll = 32646UI

            ''' <summary>
            ''' Double-pointed arrow pointing northeast and southwest.
            ''' </summary>
            Size_NESW = 32643UI

            ''' <summary>
            ''' Double-pointed arrow pointing north and south.
            ''' </summary>
            Size_NS = 32645UI

            ''' <summary>
            ''' Double-pointed arrow pointing northwest and southeast.
            ''' </summary>
            Size_NWSE = 32642UI

            ''' <summary>
            ''' Double-pointed arrow pointing west and east.
            ''' </summary>
            Size_WE = 32644UI

            ''' <summary>
            ''' Vertical arrow.
            ''' </summary>
            Up = 32516UI

            ''' <summary>
            ''' Hourglass.
            ''' </summary>
            Wait = 32514UI

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="wEventId"/> parameter of <see cref="EnvironmentUtil.NativeMethods.SHChangeNotify"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762118%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <Flags>
        Friend Enum SHChangeNotifyEventID As UInteger

            ''' <summary>
            ''' All events have occurred.
            ''' </summary>
            AllEvents = &H7FFFFFFFUI

            ''' <summary>
            ''' A folder has been created. 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that was created.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DirectoryCreated = &H8UI

            ''' <summary>
            ''' A folder has been removed.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that was removed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DirectoryDeleted = &H10UI

            ''' <summary>
            ''' The name of a folder has changed.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the previous pointer to an item identifier list (PIDL) or name of the folder.
            ''' <paramref name="dwItem2"/> contains the new PIDL or name of the folder.
            ''' </summary>
            DirectoryRenamed = &H20000UI

            ''' <summary>
            ''' A non-folder item has been created.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the item that was created.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            ItemCreated = &H2UI

            ''' <summary>
            ''' A nonfolder item has been deleted.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the item that was deleted.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            ItemDeleted = &H4UI

            ''' <summary>
            ''' The name of a nonfolder item has changed.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the previous PIDL or name of the item.
            ''' <paramref name="dwItem2"/> contains the new PIDL or name of the item.
            ''' </summary>
            ItemRenamed = &H1UI

            ''' <summary>
            ''' A drive has been added.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive that was added.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DriveAdded = &H100UI

            ''' <summary>
            ''' A drive has been added and the Shell should create a new window for the drive.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive that was added.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DriveAddedShell = &H10000UI

            ''' <summary>
            ''' A drive has been removed. 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive that was removed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DriveRemoved = &H80UI

            ''' <summary>
            ''' Storage media has been inserted into a drive.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive that contains the new media.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            MediaInserted = &H20UI

            ''' <summary>
            ''' Storage media has been removed from a drive.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive from which the media was removed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            MediaRemoved = &H40UI

            ''' <summary>
            ''' A folder on the local computer is being shared via the network.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that is being shared.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            NetShared = &H200UI

            ''' <summary>
            ''' A folder on the local computer is no longer being shared via the network.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that is no longer being shared.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            NetUnshared = &H400UI

            ''' <summary>
            ''' The computer has disconnected from a server.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the server from which the computer was disconnected.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            ServerDisconnected = &H4000UI

            ''' <summary>
            ''' The attributes of an item or folder have changed.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the item or folder that has changed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            ItemAttributesChanged = &H800UI

            ''' <summary>
            ''' A file type association has changed. 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> must be specified in the <paramref name="uFlags"/> parameter.
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are not used and must be set as <see cref="IntPtr.Zero"/>.
            ''' </summary>
            FileAssocChanged = &H8000000UI

            ''' <summary>
            ''' The amount of free space on a drive has changed.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive on which the free space changed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            FreespaceChanged = &H40000UI

            ''' <summary>
            ''' The contents of an existing folder have changed but the folder still exists and has not been renamed.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that has changed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' If a folder has been created, deleted or renamed use <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.DirectoryCreated"/>, 
            ''' or <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.DirectoryRenamed"/> respectively instead.
            ''' </summary>
            UpdateDirectory = &H1000UI

            ''' <summary>
            ''' An image in the system image list has changed.
            ''' <see cref="EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.DWORD"/> must be specified in <paramref name="uFlags"/>.
            ''' </summary>
            UpdateImage = &H8000UI

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="uFlags"/> parameter of <see cref="EnvironmentUtil.NativeMethods.SHChangeNotify"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762118%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <Flags()>
        Friend Enum SHChangeNotifyFlags

            ''' <summary>
            ''' The <paramref name="dwItem1"/> and <paramref name="dwItem2"/> parameters are DWORD values.
            ''' </summary>
            Dword = &H3

            ''' <summary>
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are the addresses of 'ITEMIDLIST' structures that
            ''' represent the item(s) affected by the change.
            ''' Each 'ITEMIDLIST' must be relative to the desktop folder.
            ''' </summary>
            ItemIdList = &H0UI

            ''' <summary>
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are the addresses of null-terminated strings, 
            ''' of maximum length MAX_PATH that contain the full path names of the items affected by the change.
            ''' </summary>
            PathA = &H1

            ''' <summary>
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are the addresses of null-terminated strings,
            ''' of maximum length MAX_PATH that contain the full path names of the items affected by the change.
            ''' </summary>
            PathW = &H5

            ''' <summary>
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are the addresses of null-terminated strings,
            ''' that represent the friendly names of the printer(s) affected by the change.
            ''' </summary>
            PrinterA = &H2

            ''' <summary>
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are the addresses of null-terminated strings,
            ''' that represent the friendly names of the printer(s) affected by the change.
            ''' </summary>
            PrinterW = &H6

            ''' <summary>
            ''' The function should not return until the notification has been delivered to all affected components.
            ''' As this flag modifies other data-type flags it cannot by used by itself.
            ''' </summary>
            Flush = &H1000

            ''' <summary>
            ''' The function should begin delivering notifications to all affected components,
            ''' but should return as soon as the notification process has begun.
            ''' As this flag modifies other data-type flags it cannot by used by itself.
            ''' </summary>
            FlushNoWait = &H2000

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="uiAction"/> parameter of <see cref="EnvironmentUtil.NativeMethods.SystemParametersInfo"/> function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947(v=vs.85).aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Enum SystemParametersActionFlags As UInteger

            ''' <summary>
            ''' Determines whether the warning beeper is on.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the beeper is on,
            ''' or <see langword="False"/> if it is off.
            ''' </summary>
            GetBeep = &H1

            ''' <summary>
            ''' Turns the warning beeper on or off. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on, or <see langword="False"/> for off.
            ''' </summary>
            SetBeep = &H2

            ''' <summary>
            ''' Retrieves the border multiplier factor that determines the width of a window's sizing border.
            ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives this value.
            ''' </summary>
            Getborder = &H5

            ''' <summary>
            ''' Sets the border multiplier factor that determines the width of a window's sizing border.
            ''' The <paramref name="uiParam"/> parameter specifies the new value.
            ''' </summary>
            SetBorder = &H6

            ''' <summary>
            ''' Retrieves the keyboard repeat-speed setting, which is a value in the range 
            ''' from 0 (approximately 2.5 repetitions per second) through 31 (approximately 30 repetitions per second). 
            ''' The actual repeat rates are hardware-dependent and may vary from a linear scale by as much as 20%. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Integer"/> variable that receives the setting
            ''' </summary>
            GetKeyboardSpeed = &HA

            ''' <summary>
            ''' Sets the keyboard repeat-speed setting. 
            ''' The <paramref name="uiParam"/> parameter must specify a value in the range 
            ''' from 0 (approximately 2.5 repetitions per second) through 31 (approximately 30 repetitions per second).
            ''' The actual repeat rates are hardware-dependent and may vary from a linear scale by as much as 20%.
            ''' If <paramref name="uiParam"/> is greater than 31, the parameter is set to 31.
            ''' </summary>
            SetKeyboardSpeed = &HB

            ''' <summary>
            ''' Sets or retrieves the width, in pixels, of an icon cell. 
            ''' The system uses this rectangle to arrange icons in large icon view.
            ''' To set this value, set <paramref name="uiParam"/> to the new value and set <paramref name="pvParam"/> to <see langword="Nothing"/>. 
            ''' You cannot set this value to less than SM_CXICON.
            ''' To retrieve this value, <paramref name="pvParam"/> must point to an <see cref="Integer"/> that receives the current value.
            ''' </summary>
            IconHorizontalSpacing = &HD

            ''' <summary>
            ''' Sets or retrieves the height, in pixels, of an icon cell.
            ''' To set this value, set <paramref name="uiParam"/> to the new value and set <paramref name="pvParam"/> to null.
            ''' You cannot set this value to less than SM_CYICON.
            ''' To retrieve this value, <paramref name="pvParam"/> must point to an <see cref="Integer"/> that receives the current value.
            ''' </summary>
            IconVerticalSpacing = &H18

            ''' <summary>
            ''' Retrieves the screen saver time-out value, in seconds. 
            ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives the value.
            ''' </summary>
            GetScreensaveTimeout = &HE

            ''' <summary>
            ''' Sets the screen saver time-out value to the value of the <paramref name="uiParam"/> parameter. 
            ''' This value is the amount of time, in seconds, that the system must be idle before the screen saver activates.
            ''' </summary>
            SetScreensaveTimeout = &HF

            ''' <summary>
            ''' Sets the state of the screen saver. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> to activate screen saving, or <see langword="False"/> to deactivate it.
            ''' </summary>
            SetScreensaveActive = &H11

            ''' <summary>
            ''' Sets the desktop wallpaper. 
            ''' The value of the <paramref name="pvParam"/> parameter determines the new wallpaper. 
            ''' To specify a wallpaper bitmap, set <paramref name="pvParam"/> to point to a null-terminated string containing the name of a bitmap file. 
            ''' Setting <paramref name="pvParam"/> to "" removes the wallpaper.
            ''' Setting <paramref name="pvParam"/> to null reverts to the default wallpaper.
            ''' </summary>
            SetDesktopWallpaper = &H14

            ''' <summary>
            ''' Retrieves the full path of the bitmap file for the desktop wallpaper.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="StringBuilder"/> that receives a null-terminated path string.
            ''' Set the <paramref name="uiParam"/> parameter to the size, in characters, of the <paramref name="pvParam"/> buffer. 
            ''' The returned string will not exceed <see cref="StringBuilder.MaxCapacity"/> characters. 
            ''' If there is no desktop wallpaper, the returned string is empty.
            ''' </summary>
            GetDesktopWallpaper = &H73

            ''' <summary>
            ''' Retrieves the keyboard repeat-delay setting, 
            ''' which is a value in the range from 0 (approximately 250 ms delay) through 3 (approximately 1 second delay). 
            ''' The actual delay associated with each value may vary depending on the hardware. 
            ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives the setting.
            ''' </summary>
            GetKeyboardDelay = &H16

            ''' <summary>
            ''' Sets the keyboard repeat-delay setting. 
            ''' The <paramref name="uiParam"/> parameter must specify 0, 1, 2, or 3, where zero sets the shortest delay
            ''' (approximately 250 ms) and 3 sets the longest delay (approximately 1 second).
            ''' The actual delay associated with each value may vary depending on the hardware.
            ''' </summary>
            SetKeyboardDelay = &H17

            ''' <summary>
            ''' Determines whether icon-title wrapping is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, 
            ''' or <see langword="False"/> otherwise.
            ''' </summary>
            GetIconTitleWrap = &H19

            ''' <summary>
            ''' Turns icon-title wrapping on or off. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on, or <see langword="False"/> for off.
            ''' </summary>
            SetIconTitleWrap = &H1A

            ''' <summary>
            ''' Determines whether pop-up menus are left-aligned or right-aligned, relative to the corresponding menu-bar item.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if left-aligned, 
            ''' or <see langword="False"/> otherwise.
            ''' </summary>
            GetMenuDropAlignment = &H1B

            ''' <summary>
            ''' Sets the alignment value of pop-up menus. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for right alignment, or <see langword="False"/> for left alignment.
            ''' </summary>
            SetMenuDropAlignment = &H1C

            ''' <summary>
            ''' Sets the width of the double-click rectangle to the value of the <paramref name="uiParam"/> parameter.
            ''' The double-click rectangle is the rectangle within which the second click of a double-click must 
            ''' fall for it to be registered as a double-click.
            ''' To retrieve the width of the double-click rectangle, call GetSystemMetrics with the SM_CXDOUBLECLK flag.
            ''' </summary>
            SetDoubleClickWidth = &H1D

            ''' <summary>
            ''' Sets the height of the double-click rectangle to the value of the <paramref name="uiParam"/> parameter.
            ''' The double-click rectangle is the rectangle within which the second click of a double-click must 
            ''' fall for it to be registered as a double-click.
            ''' To retrieve the height of the double-click rectangle, call GetSystemMetrics with the SM_CYDOUBLECLK flag.
            ''' </summary>
            SetDoubleClickHeight = &H1E

            ''' <summary>
            ''' Sets the double-click time for the mouse to the value of the <paramref name="uiParam"/> parameter. 
            ''' The double-click time is the maximum number of milliseconds that can occur between the 
            ''' first and second clicks of a double-click. 
            ''' You can also call the SetDoubleClickTime function to set the double-click time. 
            ''' To get the current double-click time, call the GetDoubleClickTime function.
            ''' </summary>
            SetDoubleclickTime = &H20

            ''' <summary>
            ''' Swaps or restores the meaning of the left and right mouse buttons. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> to swap the meanings of the buttons, or <see langword="False"/> to restore their original meanings.
            ''' </summary>
            SetMousebuttonSwap = &H21

            ''' <summary>
            ''' Determines whether dragging of full windows is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, or <see langword="False"/> otherwise.
            ''' </summary>
            GetDragFullWindows = &H26

            ''' <summary>
            ''' Sets dragging of full windows either on or off. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on, or <see langword="False"/> for off.
            ''' </summary>
            SetDragFullWindows = &H25

            ''' <summary>
            ''' Determines whether the font smoothing feature is enabled. 
            ''' This feature uses font antialiasing to make font curves appear smoother by painting pixels at different gray levels.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the feature is enabled,
            '''  or <see langword="False"/> if it is not.
            ''' Windows 95:  This flag is supported only if Windows Plus! is installed. See GETWINDOWSEXTENSION.
            ''' </summary>
            GetFontSmoothing = &H4A

            ''' <summary>
            ''' Enables or disables the font smoothing feature, which uses font antialiasing to make font curves appear smoother
            ''' by painting pixels at different gray levels.
            ''' To enable the feature, set the <paramref name="uiParam"/> parameter to TRUE. 
            ''' To disable the feature, set <paramref name="uiParam"/> to FALSE.
            ''' </summary>
            SetFontSmoothing = &H4B

            ''' <summary>
            ''' Sets the width, in pixels, of the rectangle used to detect the start of a drag operation. 
            ''' Set <paramref name="uiParam"/> to the new value.
            ''' To retrieve the drag width, call GetSystemMetrics with the SM_CXDRAG flag.
            ''' </summary>
            SetDragWidth = &H4C

            ''' <summary>
            ''' Sets the height, in pixels, of the rectangle used to detect the start of a drag operation. 
            ''' Set <paramref name="uiParam"/> to the new value.
            ''' To retrieve the drag height, call GetSystemMetrics with the SM_CYDRAG flag.
            ''' </summary>
            SetDragHeight = &H4D

            ''' <summary>
            ''' Reloads the system cursors. 
            ''' Set the <paramref name="uiParam"/> parameter to zero and the <paramref name="pvParam"/> parameter to null.
            ''' </summary>
            Setcursors = &H57

            ''' <summary>
            ''' Reloads the system icons. 
            ''' Set the <paramref name="uiParam"/> parameter to zero and the <paramref name="pvParam"/> parameter to null.
            ''' </summary>
            Seticons = &H58

            ''' <summary>
            ''' Retrieves the input locale identifier for the system default input language. 
            ''' The <paramref name="pvParam"/> parameter must point to an HKL variable that receives this value. 
            ''' For more information, see Languages, Locales, and Keyboard Layouts on MSDN.
            ''' </summary>
            GetDefaultInputLang = &H59

            ''' <summary>
            ''' Sets the default input language for the system shell and applications. 
            ''' The specified language must be displayable using the current system character set. 
            ''' The <paramref name="pvParam"/> parameter must point to an HKL variable that contains the input locale identifier for the default language. 
            ''' For more information, see Languages, Locales, and Keyboard Layouts on MSDN.
            ''' </summary>
            SetDefaultInputLang = &H5A

            ''' <summary>
            ''' Sets the hot key set for switching between input languages. 
            ''' The <paramref name="uiParam"/> and <paramref name="pvParam"/> parameters are not used.
            ''' The value sets the shortcut keys in the keyboard property sheets by reading the registry again. 
            ''' The registry must be set before this flag is used. 
            ''' the path in the registry is \HKEY_CURRENT_USER\keyboard layout\toggle. 
            ''' Valid values are "1" = ALT+SHIFT, "2" = CTRL+SHIFT, and "3" = none.
            ''' </summary>
            SetLangToggle = &H5B

            ''' <summary>
            ''' Determines whether the Mouse Trails feature is enabled. 
            ''' This feature improves the visibility of mouse cursor movements by briefly showing a trail of cursors and quickly erasing them.
            ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives a value. 
            ''' If the value is zero or 1, the feature is disabled.
            ''' If the value is greater than 1, the feature is enabled and the value indicates the number of cursors drawn in the trail.
            ''' The <paramref name="uiParam"/> parameter is not used.
            ''' </summary>
            GetMouseTrails = &H5E

            ''' <summary>
            ''' Enables or disables the Mouse Trails feature, which improves the visibility of mouse cursor movements by briefly showing
            ''' a trail of cursors and quickly erasing them.
            ''' To disable the feature, set the <paramref name="uiParam"/> parameter to zero or 1. 
            ''' To enable the feature, set <paramref name="uiParam"/> to a value greater than 1 to indicate the number of cursors drawn in the trail.
            ''' </summary>
            SetMouseTrails = &H5D

            ''' <summary>
            ''' Determines whether the snap-to-default-button feature is enabled. 
            ''' If enabled, the mouse cursor automatically moves to the default button, such as "OK" or "Apply", of a dialog box. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the feature is on, 
            ''' or <see langword="False"/> if it is off.
            ''' </summary>
            GetSnapToDefButton = &H5F

            ''' <summary>
            ''' Enables or disables the snap-to-default-button feature. 
            ''' If enabled, the mouse cursor automatically moves to the default button, such as "OK" or "Apply", of a dialog box. 
            ''' Set the <paramref name="uiParam"/> parameter to <see langword="True"/> to enable the feature, or <see langword="False"/> to disable it.
            ''' Applications should use the ShowWindow function when displaying a dialog box so the dialog manager can position the mouse cursor.
            ''' </summary>
            SetSnapToDefButton = &H60

            ''' <summary>
            ''' Retrieves the width, in pixels, of the rectangle within which the mouse pointer has to stay for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' The <paramref name="pvParam"/> parameter must point to a UINT variable that receives the width.
            ''' </summary>
            GetMouseHoverWidth = &H62

            ''' <summary>
            ''' Retrieves the width, in pixels, of the rectangle within which the mouse pointer has to stay for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' The <paramref name="pvParam"/> parameter must point to a UINT variable that receives the width.
            ''' </summary>
            SetMouseHoverWidth = &H63

            ''' <summary>
            ''' Retrieves the height, in pixels, of the rectangle within which the mouse pointer has to stay for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' The <paramref name="pvParam"/> parameter must point to a UINT variable that receives the height.
            ''' </summary>
            GetMouseHoverHeight = &H64

            ''' <summary>
            ''' Sets the height, in pixels, of the rectangle within which the mouse pointer has to stay for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' Set the <paramref name="uiParam"/> parameter to the new height.
            ''' </summary>
            SetMouseHoverHeight = &H65

            ''' <summary>
            ''' Retrieves the time, in milliseconds, that the mouse pointer has to stay in the hover rectangle for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' The <paramref name="pvParam"/> parameter must point to a UINT variable that receives the time.
            ''' </summary>
            GetMouseHoverTime = &H66

            ''' <summary>
            ''' Sets the time, in milliseconds, that the mouse pointer has to stay in the hover rectangle for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' This is used only if you pass HOVER_DEFAULT in the dwHoverTime parameter in the call to TrackMouseEvent. 
            ''' Set the <paramref name="uiParam"/> parameter to the new time.
            ''' </summary>
            SetMouseHoverTime = &H67

            ''' <summary>
            ''' Retrieves the number of lines to scroll when the mouse wheel is rotated. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UInteger"/> variable that receives the number of lines. 
            ''' The default value is 3.
            ''' </summary>
            GetWheelScrollLines = &H68

            ''' <summary>
            ''' Sets the number of lines to scroll when the mouse wheel is rotated. 
            ''' The number of lines is set from the <paramref name="uiParam"/> parameter.
            ''' The number of lines is the suggested number of lines to scroll when the mouse wheel is rolled without using modifier keys.
            ''' If the number is 0, then no scrolling should occur. 
            ''' If the number of lines to scroll is greater than the number of lines viewable,
            ''' and in particular if it is WHEEL_PAGESCROLL (#defined as UINT_MAX), the scroll operation should be interpreted
            ''' as clicking once in the page down or page up regions of the scroll bar.
            ''' </summary>
            SetWheelScrollLines = &H69

            ''' <summary>
            ''' Retrieves the time, in milliseconds, that the system waits before displaying a shortcut menu when the mouse cursor is over a submenu item. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UShort"/> variable that receives the time of the delay.
            ''' </summary>
            GetMenuShowDelay = &H6A

            ''' <summary>
            ''' Sets <paramref name="uiParam"/> to the time, in milliseconds, that the system waits before displaying a shortcut menu when the mouse cursor is
            ''' over a submenu item.
            ''' </summary>
            SetMenuShowDelay = &H6B

            ''' <summary>
            ''' Retrieves the current mouse speed. 
            ''' The mouse speed determines how far the pointer will move based on the distance the mouse moves.
            ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives a value which 
            ''' ranges between 1 (slowest) and 20 (fastest).
            ''' A value of 10 is the default. 
            ''' The value can be set by an end user using the mouse control panel application or by an application using SETMOUSESPEED.
            ''' </summary>
            GetMouseSpeed = &H70

            ''' <summary>
            ''' Sets the current mouse speed. 
            ''' The <paramref name="pvParam"/> parameter is an <see cref="Integer"/> variable between 1 (slowest) and 20 (fastest). 
            ''' A value of 10 is the default.
            ''' This value is typically set using the mouse control panel application.
            ''' </summary>
            SetMouseSpeed = &H71

            ''' <summary>
            ''' Determines whether active window tracking (activating the window the mouse is on) is on or off. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for on, 
            ''' or <see langword="False"/> for off.
            ''' </summary>
            GetActiveWindowTracking = &H1000

            ''' <summary>
            ''' Sets active window tracking (activating the window the mouse is on) either on or off. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ''' </summary>
            SetActiveWindowTracking = &H1001

            ''' <summary>
            ''' Determines whether the menu animation feature is enabled. 
            ''' This master switch must be on to enable menu animation effects.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if animation is enabled 
            ''' and <see langword="False"/> if it is disabled.
            ''' If animation is enabled, GETMENUFADE indicates whether menus use fade or slide animation.
            ''' </summary>
            GetMenuAnimation = &H1002

            ''' <summary>
            ''' Enables or disables menu animation. 
            ''' This master switch must be on for any menu animation to occur.
            ''' The <paramref name="pvParam"/> parameter is a <see cref="Boolean"/> variable; 
            ''' set <paramref name="pvParam"/> to <see langword="True"/> to enable animation and <see langword="False"/> to disable animation.
            ''' If animation is enabled, GETMENUFADE indicates whether menus use fade or slide animation.
            ''' </summary>
            SetMenuAnimation = &H1003

            ''' <summary>
            ''' Determines whether the slide-open effect for combo boxes is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for enabled, 
            ''' or <see langword="False"/> for disabled.
            ''' </summary>
            GetComboboxAnimation = &H1004

            ''' <summary>
            ''' Enables or disables the slide-open effect for combo boxes. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable the gradient effect, or <see langword="False"/> to disable it.
            ''' </summary>
            SetComboboxAnimation = &H1005

            ''' <summary>
            ''' Determines whether the smooth-scrolling effect for list boxes is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point toa <see cref="Boolean"/> variable that receives <see langword="True"/> for enabled, 
            ''' or <see langword="False"/> for disabled.
            ''' </summary>
            GetListboxSmoothScrolling = &H1006

            ''' <summary>
            ''' Enables or disables the smooth-scrolling effect for list boxes. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable the smooth-scrolling effect,
            ''' or <see langword="False"/> to disable it.
            ''' </summary>
            SetListboxSmoothScrolling = &H1007

            ''' <summary>
            ''' Determines whether the gradient effect for window title bars is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for enabled, 
            ''' or <see langword="False"/> for disabled. 
            ''' For more information about the gradient effect, see the GetSysColor function.
            ''' </summary>
            GetGradientCaptions = &H1008

            ''' <summary>
            ''' Enables or disables the gradient effect for window title bars. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable it, or <see langword="False"/> to disable it.
            ''' The gradient effect is possible only if the system has a color depth of more than 256 colors. For more information about
            ''' the gradient effect, see the GetSysColor function.
            ''' </summary>
            SetGradientCaptions = &H1009

            ''' <summary>
            ''' Determines whether menu access keys are always underlined. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that 
            ''' receives <see langword="True"/> if menu access keys are always underlined, 
            ''' and <see langword="False"/> if they are underlined only when the menu is activated by the keyboard.
            ''' </summary>
            GetKeyboardCues = &H100A

            ''' <summary>
            ''' Sets the underlining of menu access key letters. 
            ''' The <paramref name="pvParam"/> parameter is a <see cref="Boolean"/> variable. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to always underline menu access keys, 
            ''' or <see langword="False"/> to underline menu access keys only when the menu is activated from the keyboard.
            ''' </summary>
            SetKeyboardCues = &H100B

            ''' <summary>
            ''' Determines whether hot tracking of user-interface elements, such as menu names on menu bars, is enabled.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for enabled, 
            ''' or <see langword="False"/> for disabled.
            ''' Hot tracking means that when the cursor moves over an item, it is highlighted but not selected. 
            ''' You can query this value to decide whether to use hot tracking in the user interface of your application.
            ''' </summary>
            GetHotTracking = &H100E

            ''' <summary>
            ''' Enables or disables hot tracking of user-interface elements such as menu names on menu bars. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable it, or <see langword="False"/> to disable it.
            ''' Hot-tracking means that when the cursor moves over an item, it is highlighted but not selected.
            ''' </summary>
            SetHotTracking = &H100F

            ''' <summary>
            ''' Determines whether menu fade animation is enabled.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/>
            ''' when fade animation is enabled and <see langword="False"/> when it is disabled. 
            ''' If fade animation is disabled, menus use slide animation.
            ''' This flag is ignored unless menu animation is enabled, which you can do using the SETMENUANIMATION flag.
            ''' For more information, see AnimateWindow.
            ''' </summary>
            GetMenuFade = &H1012

            ''' <summary>
            ''' Enables or disables menu fade animation. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable the menu fade effect or <see langword="False"/> to disable it.
            ''' If fade animation is disabled, menus use slide animation.
            ''' The menu fade effect is possible only if the system has a color depth of more than 256 colors. 
            ''' This flag is ignored unless MENUANIMATION is also set. 
            ''' For more information, see AnimateWindow.
            ''' </summary>
            SetMenuFade = &H1013

            ''' <summary>
            ''' Determines whether the selection fade effect is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled 
            ''' or <see langword="False"/> if disabled.
            ''' The selection fade effect causes the menu item selected by the user to remain on the screen briefly while fading out
            ''' after the menu is dismissed.
            ''' </summary>
            GetSelectionFade = &H1014

            ''' <summary>
            ''' The selection fade effect causes the menu item selected by the user to remain on the screen briefly while fading out
            ''' after the menu is dismissed. 
            ''' The selection fade effect is possible only if the system has a color depth of more than 256 colors.
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable the selection fade effect or <see langword="False"/> to disable it.
            ''' </summary>
            SetSelectionFade = &H1015

            ''' <summary>
            ''' Determines whether ToolTip animation is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled 
            ''' or <see langword="False"/> if disabled. 
            ''' If ToolTip animation is enabled, GETTOOLTIPFADE indicates whether ToolTips use fade or slide animation.
            ''' </summary>
            GetTooltipAnimation = &H1016

            ''' <summary>
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable ToolTip animation or <see langword="False"/> to disable it. 
            ''' If enabled, you can use SETTOOLTIPFADE to specify fade or slide animation.
            ''' </summary>
            SetTooltipAnimation = &H1017

            ''' <summary>
            ''' Determines whether the cursor has a shadow around it. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the shadow is enabled, 
            ''' <see langword="False"/> if it is disabled. 
            ''' This effect appears only if the system has a color depth of more than 256 colors.
            ''' </summary>
            GetCursorShadow = &H101A

            ''' <summary>
            ''' Enables or disables a shadow around the cursor. 
            ''' The <paramref name="pvParam"/> parameter is a <see cref="Boolean"/> variable. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable the shadow or <see langword="False"/> to disable the shadow. 
            ''' This effect appears only if the system has a color depth of more than 256 colors.
            ''' </summary>
            SetCursorShadow = &H101B

            ''' <summary>
            ''' Retrieves the state of the Mouse Sonar feature. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled or
            ''' <see langword="False"/> otherwise. 
            ''' For more information, see About Mouse Input on MSDN.
            ''' </summary>
            GetMouseSonar = &H101C

            ''' <summary>
            ''' Turns the Sonar accessibility feature on or off. 
            ''' This feature briefly shows several concentric circles around the mouse pointer when the user presses and releases the CTRL key. 
            ''' The <paramref name="pvParam"/> parameter specifies <see langword="True"/> for on and <see langword="False"/> for off. 
            ''' The default is off.
            ''' For more information, see About Mouse Input.
            ''' </summary>
            SetMouseSonar = &H101D

            ''' <summary>
            ''' Retrieves the state of the Mouse ClickLock feature. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, 
            ''' or <see langword="False"/> otherwise. 
            ''' For more information, see About Mouse Input.
            ''' </summary>
            GetMouseClickLock = &H101E

            ''' <summary>
            ''' Turns the Mouse ClickLock accessibility feature on or off. 
            ''' This feature temporarily locks down the primary mouse button when that button is clicked and 
            ''' held down for the time specified by SETMOUSECLICKLOCKTIME. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on, or <see langword="False"/> for off. The default is off. 
            ''' For more information, see Remarks and About Mouse Input on MSDN.
            ''' </summary>
            SetMouseClickLock = &H101F

            ''' <summary>
            ''' Retrieves the state of the Mouse Vanish feature. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled or 
            ''' <see langword="False"/> otherwise. 
            ''' For more information, see About Mouse Input on MSDN.
            ''' </summary>
            GetMouseVanish = &H1020

            ''' <summary>
            ''' Turns the Vanish feature on or off. 
            ''' This feature hides the mouse pointer when the user types; the pointer reappears when the user moves the mouse. 
            ''' The <paramref name="pvParam"/> parameter specifies <see langword="True"/> for on and <see langword="False"/> for off. The default is off.
            ''' For more information, see About Mouse Input on MSDN.
            ''' </summary>
            SetMouseVanish = &H1021

            ''' <summary>
            ''' Determines whether native User menus have flat menu appearance. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that returns <see langword="True"/> if the 
            ''' flat menu appearance is set, or <see langword="False"/> otherwise.
            ''' </summary>
            GetFlatMenu = &H1022

            ''' <summary>
            ''' Enables or disables flat menu appearance for native User menus. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable flat menu appearance or <see langword="False"/> to disable it.
            ''' When enabled, the menu bar uses COLOR_MENUBAR for the menubar background, COLOR_MENU for the menu-popup background, COLOR_MENUHILIGHT
            ''' for the fill of the current menu selection, and COLOR_HILIGHT for the outline of the current menu selection.
            ''' If disabled, menus are drawn using the same metrics and colors as in Windows 2000 and earlier.
            ''' </summary>
            SetFlatMenu = &H1023

            ''' <summary>
            ''' Determines whether the drop shadow effect is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that returns <see langword="True"/> if enabled or 
            ''' <see langword="False"/> if disabled.
            ''' </summary>
            GetDropShadow = &H1024

            ''' <summary>
            ''' Enables or disables the drop shadow effect.
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable the drop shadow effect or <see langword="False"/> to disable it.
            ''' You must also have CS_DROPSHADOW in the window class style.
            ''' </summary>
            SetDropShadow = &H1025

            ''' <summary>
            ''' Retrieves a <see cref="Boolean"/> indicating whether an application can reset the screensaver's timer by calling the SendInput function
            ''' to simulate keyboard or mouse input. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/>
            ''' if the simulated input will be blocked, or <see langword="False"/> otherwise.
            ''' </summary>
            GetBlockSendInputResets = &H1026

            ''' <summary>
            ''' Determines whether an application can reset the screensaver's timer by calling the SendInput function to simulate keyboard
            ''' or mouse input. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> if the screensaver will not be deactivated by simulated input,
            ''' or <see langword="False"/> if the screensaver will be deactivated by simulated input.
            ''' </summary>
            SetBlockSendInputResets = &H1027

            ''' <summary>
            ''' Determines whether UI effects are enabled or disabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/>
            ''' if all UI effects are enabled, or <see langword="False"/> if they are disabled.
            ''' </summary>
            GetUiEffects = &H103E

            ''' <summary>
            ''' Enables or disables UI effects. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable all UI effects or <see langword="False"/> to disable all UI effects.
            ''' </summary>
            SetUiEffects = &H103F

            ''' <summary>
            ''' Retrieves the amount of time following user input, in milliseconds, during which the system will not allow applications
            ''' to force themselves into the foreground. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UShort"/> variable that receives the time.
            ''' </summary>
            GetForegroundLockTimeout = &H2000

            ''' <summary>
            ''' Sets the amount of time following user input, in milliseconds, during which the system does not allow applications
            ''' to force themselves into the foreground. 
            ''' Set <paramref name="pvParam"/> to the new timeout value.
            ''' The calling thread must be able to change the foreground window, otherwise the call fails.
            ''' </summary>
            SetForegroundLockTimeout = &H2001

            ''' <summary>
            ''' Retrieves the active window tracking delay, in milliseconds. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UShort"/> variable that receives the time.
            ''' </summary>
            GetActiveWndTrkTimeout = &H2002

            ''' <summary>
            ''' Sets the active window tracking delay. 
            ''' Set <paramref name="pvParam"/> to the number of milliseconds to delay before activating the window under the mouse pointer.
            ''' </summary>
            SetActiveWndTrkTimeout = &H2003

            ''' <summary>
            ''' Retrieves the number of times SetForegroundWindow will flash the taskbar button when rejecting a foreground switch request.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UShort"/> variable that receives the value.
            ''' </summary>
            GetForegroundFlashCount = &H2004

            ''' <summary>
            ''' Sets the number of times SetForegroundWindow will flash the taskbar button when rejecting a foreground switch request.
            ''' Set <paramref name="pvParam"/> to the number of times to flash.
            ''' </summary>
            SetForegroundFlashCount = &H2005

            ''' <summary>
            ''' Retrieves the caret width in edit controls, in pixels. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UShort"/> that receives this value.
            ''' </summary>
            GetCaretWidth = &H2006

            ''' <summary>
            ''' Sets the caret width in edit controls. 
            ''' Set <paramref name="pvParam"/> to the desired width, in pixels. 
            ''' The default and minimum value is 1.
            ''' </summary>
            SetCaretWidth = &H2007

            ''' <summary>
            ''' Retrieves the time delay before the primary mouse button is locked. 
            ''' The <paramref name="pvParam"/> parameter must point to <see cref="Integer"/> that receives the time delay. 
            ''' This is only enabled if SETMOUSECLICKLOCK is set to <see langword="True"/>. 
            ''' For more information, see About Mouse Input on MSDN.
            ''' </summary>
            GetMouseClickLockTime = &H2008

            ''' <summary>
            ''' Turns the Mouse ClickLock accessibility feature on or off. 
            ''' This feature temporarily locks down the primary mouse button when that button is clicked and 
            ''' held down for the time specified by SETMOUSECLICKLOCKTIME. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on, or <see langword="False"/> for off. 
            ''' The default is off. 
            ''' For more information, see Remarks and About Mouse Input on MSDN.
            ''' </summary>
            SetMouseClickLockTime = &H2009

            ''' <summary>
            ''' Retrieves a contrast value that is used in ClearType™ smoothing. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UInteger"/> that receives the information.
            ''' </summary>
            GetFontSmoothingContrast = &H200C

            ''' <summary>
            ''' Sets the contrast value used in ClearType smoothing. 
            ''' The <paramref name="pvParam"/> parameter points to a <see cref="UInteger"/> that holds the contrast value.
            ''' Valid contrast values are from 1000 to 2200. The default value is 1400.
            ''' When using this option, the fWinIni parameter must be set to SPIF_SENDWININICHANGE | SPIF_UPDATEINIFILE; otherwise,
            ''' SystemParametersInfo fails.
            ''' SETFONTSMOOTHINGTYPE must also be set to FE_FONTSMOOTHINGCLEARTYPE.
            ''' </summary>
            SetFontSmoothingContrast = &H200D

            ''' <summary>
            ''' Retrieves the width, in pixels, of the left and right edges of the focus rectangle drawn with DrawFocusRect.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UInteger"/>.
            ''' </summary>
            GetFocusBorderWidth = &H200E

            ''' <summary>
            ''' Sets the height of the left and right edges of the focus rectangle drawn with DrawFocusRect to the value of 
            ''' the <paramref name="pvParam"/> parameter.
            ''' </summary>
            SetFocusBorderWidth = &H200F

            ''' <summary>
            ''' Retrieves the height, in pixels, of the top and bottom edges of the focus rectangle drawn with DrawFocusRect.
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UInteger"/>.
            ''' </summary>
            GetFocusBorderHeight = &H2010

            ''' <summary>
            ''' Sets the height of the top and bottom edges of the focus rectangle drawn with DrawFocusRect to the 
            ''' value of the <paramref name="pvParam"/> parameter.
            ''' </summary>
            SetFocusBorderHeight = &H2011

            ''' <summary>
            ''' Determines whether animations are enabled or disabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if animations are enabled, 
            ''' or <see langword="False"/> otherwise.
            ''' </summary>
            GetClientAreaAnimation = &H1042

            ''' <summary>
            ''' Turns client area animations on or off. 
            ''' The <paramref name="pvParam"/> parameter is a <see cref="Boolean"/> variable. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable animations and other transient effects in the client area, 
            ''' or <see langword="False"/> to disable them..
            ''' </summary>
            SetClientAreaAnimation = &H1043

            ''' <summary>
            ''' Determines whether overlapped content is enabled or disabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, 
            ''' or <see langword="False"/> otherwise.
            ''' </summary>
            GetDisableOverlappedContent = &H1040

            ''' <summary>
            ''' Turns overlapped content (such as background images and watermarks) on or off. 
            ''' The <paramref name="pvParam"/> parameter is a <see cref="Boolean"/> variable. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to disable overlapped content, or <see langword="False"/> to enable overlapped content.
            ''' </summary>
            SetDisableOverlappedContent = &H1041

            ''' <summary>
            ''' Retrieves the time that notification pop-ups should be displayed, in seconds. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="ULong"/> that receives the message duration.
            ''' </summary>
            GetMessageDuration = &H2016

            ''' <summary>
            ''' Sets the time that notification pop-ups should be displayed, in seconds. 
            ''' The <paramref name="pvParam"/> parameter specifies the message duration.
            ''' </summary>
            SetMessageDuration = &H2017

            ''' <summary>
            ''' Determines whether ClearType is enabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if ClearType is enabled, or <see langword="False"/> otherwise.
            ''' </summary>
            GetCleartype = &H1048

            ''' <summary>
            ''' Turns ClearType on or off. 
            ''' The <paramref name="pvParam"/> parameter is a <see cref="Boolean"/> variable. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable ClearType, or <see langword="False"/> to disable it.
            ''' </summary>
            SetCleartype = &H1049

            ''' <summary>
            ''' Starting with Windows 8: Determines whether the system language bar is enabled or disabled. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the language bar is enabled, or <see langword="False"/> otherwise.
            ''' </summary>
            GetSystemlanguageBar = &H1050

            ''' <summary>
            ''' Starting with Windows 8: Turns the legacy language bar feature on or off. 
            ''' The <paramref name="pvParam"/> parameter is a pointer to a <see cref="Boolean"/> variable. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable the legacy language bar, or <see langword="False"/> to disable it. 
            ''' The flag is supported on Windows 8 where the legacy language bar is replaced by Input Switcher and therefore turned off by default. 
            ''' Turning the legacy language bar on is provided for compatibility reasons and has no effect on the Input Switcher.
            ''' </summary>
            SetSystemlanguageBar = &H1051

            ''' <summary>
            ''' Retrieves the number of characters to scroll when the horizontal mouse wheel is moved. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UInteger"/> variable that receives the number of lines. 
            ''' The default value is 3.
            ''' </summary>
            GetWheelscrollChars = &H6C

            ''' <summary>
            ''' Sets the number of characters to scroll when the horizontal mouse wheel is moved. 
            ''' The number of characters is set from the <paramref name="uiParam"/> parameter.
            ''' </summary>
            SetWheelscrollChars = &H6D

            ''' <summary>
            ''' Determines whether the screen saver requires a password to display the Windows desktop. 
            ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the screen saver requires a password, or <see langword="False"/> otherwise. 
            ''' The <paramref name="uiParam"/> parameter is ignored.
            ''' </summary>
            GetScreensaveSecure = &H76

            ''' <summary>
            ''' Sets whether the screen saver requires the user to enter a password to display the Windows desktop. 
            ''' The <paramref name="uiParam"/> parameter is a <see cref="Boolean"/> variable. 
            ''' The <paramref name="pvParam"/> parameter is ignored. 
            ''' Set <paramref name="uiParam"/> to <see langword="True"/> to require a password, or <see langword="False"/> to not require a password.
            ''' If the machine has entered power saving mode or system lock state, an ERROR_OPERATION_IN_PROGRESS exception occurs.
            ''' </summary>
            SetScreensaveSecure = &H77

            ''' <summary>
            ''' Retrieves the number of milliseconds that a thread can go without dispatching a message before the system considers it unresponsive. 
            ''' The <paramref name="pvParam"/> parameter must point to an integer variable that receives the value.
            ''' </summary>
            GetHungAppTimeout = &H78

            ''' <summary>
            ''' Sets the hung application time-out to the value of the <paramref name="uiParam"/> parameter. 
            ''' This value is the number of milliseconds that a thread can go without dispatching a message before the system considers it unresponsive
            ''' </summary>
            SetHungAppTimeout = &H79

            ''' <summary>
            ''' Retrieves the number of milliseconds that the system waits before terminating an application that does not respond to a shutdown request. 
            ''' The <paramref name="pvParam"/> parameter must point to an integer variable that receives the value.
            ''' </summary>
            GetWaitToKillTimeout = &H7A

            ''' <summary>
            ''' Sets the application shutdown request time-out to the value of the <paramref name="uiParam"/> parameter. 
            ''' This value is the number of milliseconds that the system waits before terminating an application that does not respond to a shutdown request
            ''' </summary>
            SetWaitToKillTimeout = &H7B

            ''' <summary>
            ''' Retrieves the number of milliseconds that the service control manager waits before 
            ''' terminating a service that does not respond to a shutdown request. 
            ''' The <paramref name="pvParam"/> parameter must point to an integer variable that receives the value.
            ''' </summary>
            GetWaitToKillServiceTimeout = &H7C

            ''' <summary>
            ''' Sets the service shutdown request time-out to the value of the <paramref name="uiParam"/> parameter. 
            ''' This value is the number of milliseconds that the system waits before terminating a service that does not respond to a shutdown request
            ''' </summary>
            SetWaitToKillServiceTimeout = &H7D

            ' ''' <summary>
            ' ''' Windows Me/98/95: Pen windows is being loaded or unloaded. 
            ' ''' The <paramref name="uiParam"/> parameter is <see langword="True"/> when loading and <see langword="False"/> when unloading pen windows. 
            ' ''' The <paramref name="pvParam"/> parameter is null.
            ' ''' </summary>
            ' SetPenWindows = &H31

            ' ''' <summary>
            ' ''' Determines whether a screen saver is currently running on the window station of the calling process.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if a screen saver is currently running, 
            ' ''' or <see langword="False"/> otherwise.
            ' ''' Note that only the interactive window station, "WinSta0", can have a screen saver running.
            ' ''' </summary>
            ' GetScreensaverRunning = &H72

            ' ''' <summary>
            ' ''' Windows Me/98:  Used internally; applications should not use this flag.
            ' ''' </summary>
            ' SetScreensaverRunning = &H61

            ' ''' <summary>
            ' ''' Retrieves information about the FilterKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a FILTERKEYS structure that receives the information. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(FILTERKEYS).
            ' ''' </summary>
            ' GetFilterKeys = &H32

            ' ''' <summary>
            ' ''' Sets the parameters of the FilterKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a FILTERKEYS structure that contains the new parameters. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(FILTERKEYS).
            ' ''' </summary>
            ' SetFilterKeys = &H33

            ' ''' <summary>
            ' ''' Retrieves information about the ToggleKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a TOGGLEKEYS structure that receives the information. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(TOGGLEKEYS).
            ' ''' </summary>
            ' GetToggleKeys = &H34

            ' ''' <summary>
            ' ''' Sets the parameters of the ToggleKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a TOGGLEKEYS structure that contains the new parameters. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(TOGGLEKEYS).
            ' ''' </summary>
            ' SetToggleKeys = &H35

            ' ''' <summary>
            ' ''' Retrieves information about the MouseKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a MOUSEKEYS structure that receives the information. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(MOUSEKEYS).
            ' ''' </summary>
            ' GetMouseKeys = &H36

            ' ''' <summary>
            ' ''' Sets the parameters of the MouseKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a MOUSEKEYS structure that contains the new parameters. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(MOUSEKEYS).
            ' ''' </summary>
            ' SetMouseKeys = &H37

            ' ''' <summary>
            ' ''' Determines whether the Show Sounds accessibility flag is on or off.
            ' ''' If it is on, the user requires an application to present information visually in situations where 
            ' ''' it would otherwise present the information only in audible form.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the feature is on, 
            ' ''' or <see langword="False"/> if it is off.
            ' ''' Using this value is equivalent to calling GetSystemMetrics (SM_SHOWSOUNDS). That is the recommended call.
            ' ''' </summary>
            ' GetShowSounds = &H38

            ' ''' <summary>
            ' ''' Sets the parameters of the SoundSentry accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a SOUNDSENTRY structure that contains the new parameters. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(SOUNDSENTRY).
            ' ''' </summary>
            ' SetShowSounds = &H39

            ' ''' <summary>
            ' ''' Retrieves information about the StickyKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a STICKYKEYS structure that receives the information. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(STICKYKEYS).
            ' ''' </summary>
            ' GetStickyKeys = &H3A

            ' ''' <summary>
            ' ''' Sets the parameters of the StickyKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a STICKYKEYS structure that contains the new parameters. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(STICKYKEYS).
            ' ''' </summary>
            ' SetStickyKeys = &H3B

            ' ''' <summary>
            ' ''' Retrieves information about the time-out period associated with the accessibility features. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an ACCESSTIMEOUT structure that receives the information. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(ACCESSTIMEOUT).
            ' ''' </summary>
            ' GetAccessTimeout = &H3C

            ' ''' <summary>
            ' ''' Sets the time-out period associated with the accessibility features. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an ACCESSTIMEOUT structure that contains the new parameters. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(ACCESSTIMEOUT).
            ' ''' </summary>
            ' SetAccessTimeout = &H3D

            ' ''' <summary>
            ' ''' Windows Me/98/95:  Retrieves information about the SerialKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a SERIALKEYS structure that receives the information. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(SERIALKEYS).
            ' ''' </summary>
            ' GetSerialKeys = &H3E

            ' ''' <summary>
            ' ''' Windows Me/98/95:  Sets the parameters of the SerialKeys accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a SERIALKEYS structure that contains the new parameters. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(SERIALKEYS).
            ' ''' </summary>
            ' SetSerialKeys = &H3F

            ' ''' <summary>
            ' ''' Retrieves information about the SoundSentry accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a SOUNDSENTRY structure that receives the information. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(SOUNDSENTRY).
            ' ''' </summary>
            ' GetSoundsEntry = &H40

            ' ''' <summary>
            ' ''' Sets the parameters of the SoundSentry accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a SOUNDSENTRY structure that contains the new parameters. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(SOUNDSENTRY).
            ' ''' </summary>
            ' SetSoundsEntry = &H41

            ' ''' <summary>
            ' ''' Determines whether audio descriptions are enabled or disabled. 
            ' ''' The <paramref name="pvParam"/> parameter is a pointer to an AUDIODESCRIPTION structure. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(AUDIODESCRIPTION).
            ' ''' </summary>
            ' GetAudioDescription = &H74

            ' ''' <summary>
            ' ''' Turns the audio descriptions feature on or off. 
            ' ''' The pvParam parameter is a pointer to an AUDIODESCRIPTION structure.
            ' ''' </summary>
            ' SetAudioDescription = &H75

            ' ''' <summary>
            ' ''' Not implemented.
            ' ''' </summary>
            ' GetFontSmoothingOrientation = &H2012

            ' ''' <summary>
            ' ''' Not implemented.
            ' ''' </summary>
            ' SetFontSmoothingOrientation = &H2013

            ' ''' <summary>
            ' ''' Retrieves the type of font smoothing. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UInteger"/> that receives the information.
            ' ''' </summary>
            ' GetFontSmoothingType = &H200A

            ' ''' <summary>
            ' ''' Sets the font smoothing type. 
            ' ''' The <paramref name="pvParam"/> parameter points to a UINT that contains either FE_FONTSMOOTHINGSTANDARD,
            ' ''' if standard anti-aliasing is used, or FE_FONTSMOOTHINGCLEARTYPE, if ClearType is used. The default is FE_FONTSMOOTHINGSTANDARD.
            ' ''' When using this option, the fWinIni parameter must be set to SPIF_SENDWININICHANGE | SPIF_UPDATEINIFILE; otherwise,
            ' ''' SystemParametersInfo fails.
            ' ''' </summary>
            ' SetFontSmoothingType = &H200B

            ' ''' <summary>
            ' ''' If SETTOOLTIPANIMATION is enabled, GETTOOLTIPFADE indicates whether ToolTip animation uses a fade effect or a slide effect.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for fade animation 
            ' ''' or <see langword="False"/> for slide animation.
            ' ''' For more information on slide and fade effects, see AnimateWindow.
            ' ''' </summary>
            ' GetTooltipFade = &H1018

            ' ''' <summary>
            ' ''' If the SETTOOLTIPANIMATION flag is enabled, use SETTOOLTIPFADE to indicate whether ToolTip animation uses a fade effect or a slide effect. 
            ' ''' Set <paramref name="pvParam"/> to <see langword="True"/> for fade animation or <see langword="False"/> for slide animation. 
            ' ''' The tooltip fade effect is possible only if the system has a color depth of more than 256 colors. 
            ' ''' For more information on the slide and fade effects, see the AnimateWindow function.
            ' ''' </summary>
            ' SetTooltipFade = &H1019

            ' ''' <summary>
            ' ''' Same as <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetKeyboardCues"/>.
            ' ''' </summary>
            ' GetMenuUnderlines = GETKEYBOARDCUES

            ' ''' <summary>
            ' ''' Same as <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetKeyboardCues"/>.
            ' ''' </summary>
            ' SetMenuUnderlines = SETKEYBOARDCUES

            ' ''' <summary>
            ' ''' Determines whether windows activated through active window tracking will be brought to the top. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for on,
            ' ''' or <see langword="False"/> for off.
            ' ''' </summary>
            ' GetActiveWndTrkZorder = &H100C

            ' ''' <summary>
            ' ''' Determines whether or not windows activated through active window tracking should be brought to the top. 
            ' ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ' ''' </summary>
            ' SetActiveWndTrkZorder = &H100D

            ' ''' <summary>
            ' ''' Determines whether the IME status window is visible (on a per-user basis). 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable  that 
            ' ''' receives <see langword="True"/> if the status window is visible, or <see langword="False"/> if it is not.
            ' ''' </summary>
            ' GetShowIMEui = &H6E

            ' ''' <summary>
            ' ''' Sets whether the IME status window is visible or not on a per-user basis. 
            ' ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on or <see langword="False"/> for off.
            ' ''' </summary>
            ' SetShowIMEui = &H6F

            ' ''' <summary>
            ' ''' Windows 95:  Determines whether the Windows extension, Windows Plus!, is installed.
            ' ''' Set the <paramref name="uiParam"/> parameter to 1.
            ' ''' The <paramref name="pvParam"/> parameter is not used. 
            ' ''' The function returns <see langword="True"/> if the extension is installed, or <see langword="False"/> if it is not.
            ' ''' </summary>
            ' GetWindowsExtension = &H5C

            ' ''' <summary>
            ' ''' Used internally; applications should not use this value.
            ' ''' </summary>
            ' SetHandheld = &H4E

            ' ''' <summary>
            ' ''' Retrieves the time-out value for the low-power phase of screen saving. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives the value. 
            ' ''' This flag is supported for 32-bit applications only.
            ' ''' </summary>
            ' GetLowPowerTimeout = &H4F

            ' ''' <summary>
            ' ''' Sets the time-out value, in seconds, for the low-power phase of screen saving. 
            ' ''' The <paramref name="uiParam"/> parameter specifies the new value.
            ' ''' The <paramref name="pvParam"/> parameter must be null. 
            ' ''' This flag is supported for 32-bit applications only.
            ' ''' </summary>
            ' SetLowPowerTimeout = &H51

            ' ''' <summary>
            ' ''' Retrieves the time-out value for the power-off phase of screen saving. The <paramref name="pvParam"/> parameter must 
            ' ''' point to an <see cref="Integer"/> variable that receives the value. 
            ' ''' This flag is supported for 32-bit applications only.
            ' ''' </summary>
            ' GetPowerOffTimeout = &H50

            ' ''' <summary>
            ' ''' Sets the time-out value, in seconds, for the power-off phase of screen saving. The <paramref name="uiParam"/> parameter specifies the new value.
            ' ''' The <paramref name="pvParam"/> parameter must be null. 
            ' ''' This flag is supported for 32-bit applications only.
            ' ''' </summary>
            ' SetPowerOffTimeout = &H52

            ' ''' <summary>
            ' ''' Determines whether the low-power phase of screen saving is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable
            ' ''' that receives <see langword="True"/> if enabled, or <see langword="False"/> if disabled. 
            ' ''' This flag is supported for 32-bit applications only.
            ' ''' </summary>
            ' GetLowPowerActive = &H53

            ' ''' <summary>
            ' ''' Activates or deactivates the low-power phase of screen saving. 
            ' ''' Set <paramref name="uiParam"/> to 1 to activate, or zero to deactivate.
            ' ''' The <paramref name="pvParam"/> parameter must be null. 
            ' ''' This flag is supported for 32-bit applications only.
            ' ''' </summary>
            ' SetLowPowerActive = &H55

            ' ''' <summary>
            ' ''' Determines whether the power-off phase of screen saving is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable
            ' ''' that receives <see langword="True"/> if enabled, or <see langword="False"/> if disabled. 
            ' ''' This flag is supported for 32-bit applications only.
            ' ''' </summary>
            ' GetPowerOffActive = &H54

            ' ''' <summary>
            ' ''' Activates or deactivates the power-off phase of screen saving. Set <paramref name="uiParam"/> to 1 to activate, or zero to deactivate.
            ' ''' The <paramref name="pvParam"/> parameter must be null. 
            ' ''' This flag is supported for 32-bit applications only.
            ' ''' </summary>
            ' SetPowerOffActive = &H56

            ' ''' <summary>
            ' ''' Retrieves information about the HighContrast accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a HIGHCONTRAST structure that receives the information. 
            ' ''' Set the cbSize member of this structure and the <paramref name="uiParam"/> parameter to sizeof(HIGHCONTRAST).
            ' ''' For a general discussion, see remarks.
            ' ''' Windows NT:  This value is not supported.
            ' ''' </summary>
            ' ''' <remarks>
            ' ''' There is a difference between the High Contrast color scheme and the High Contrast Mode. 
            ' ''' The High Contrast color scheme changes the system colors to colors that have obvious contrast;
            ' ''' you switch to this color scheme by using the Display Options in the control panel.
            ' ''' The High Contrast Mode, which uses GETHIGHCONTRAST and SETHIGHCONTRAST, advises applications to 
            ' ''' modify their appearance for visually-impaired users. 
            ' ''' It involves such things as audible warning to users and customized color scheme
            ' ''' (using the Accessibility Options in the control panel). 
            ' ''' For more information, see HIGHCONTRAST on MSDN.
            ' ''' For more information on general accessibility features, see Accessibility on MSDN.
            ' ''' </remarks>
            ' GetHighContrast = &H42

            ' ''' <summary>
            ' ''' Sets the parameters of the HighContrast accessibility feature. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a HIGHCONTRAST structure that contains the new parameters. 
            ' ''' Set the cbSize member of this structure and the <paramref name="uiParam"/> parameter to sizeof(HIGHCONTRAST).
            ' ''' Windows NT: This value is not supported.
            ' ''' </summary>
            ' SetHighContrast = &H43

            ' ''' <summary>
            ' ''' Determines whether the user relies on the keyboard instead of the mouse, 
            ' ''' and wants applications to display keyboard interfaces that would otherwise be hidden. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the user relies on the keyboard; 
            ' ''' or <see langword="False"/> otherwise.
            ' ''' Windows NT: This value is not supported.
            ' ''' </summary>
            ' GetKeyboardPref = &H44

            ' ''' <summary>
            ' ''' Sets the keyboard preference. 
            ' ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> if the user relies on the keyboard instead of the mouse,
            ' ''' and wants applications to display keyboard interfaces that would otherwise be hidden; <paramref name="uiParam"/> is <see langword="False"/> otherwise.
            ' ''' Windows NT: This value is not supported.
            ' ''' </summary>
            ' SetKeyboardPref = &H45

            ' ''' <summary>
            ' ''' Determines whether a screen reviewer utility is running. A screen reviewer utility directs textual information to an output device,
            ' ''' such as a speech synthesizer or Braille display. When this flag is set, an application should provide textual information
            ' ''' in situations where it would otherwise present the information graphically.
            ' ''' The <paramref name="pvParam"/> parameter is a pointer to a <see cref="Boolean"/> variable that receives <see langword="True"/> if a screen reviewer utility is running, or <see langword="False"/> otherwise.
            ' ''' Windows NT:  This value is not supported.
            ' ''' </summary>
            ' GetScreenReader = &H46

            ' ''' <summary>
            ' ''' Determines whether a screen review utility is running. The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on, or <see langword="False"/> for off.
            ' ''' Windows NT:  This value is not supported.
            ' ''' </summary>
            ' SetScreenReader = &H47

            ' ''' <summary>
            ' ''' Retrieves the animation effects associated with user actions. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an ANIMATIONINFO structure that receives the information. 
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(ANIMATIONINFO).
            ' ''' </summary>
            ' GetAnimation = &H48

            ' ''' <summary>
            ' ''' Sets the animation effects associated with user actions. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an ANIMATIONINFO structure that contains the new parameters.
            ' ''' Set the <paramref name="cbSize"/> member of this structure and the <paramref name="uiParam"/> parameter to sizeof(ANIMATIONINFO).
            ' ''' </summary>
            ' SetAnimation = &H49

            ' ''' <summary>
            ' ''' Retrieves the metrics associated with the nonclient area of nonminimized windows. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a NONCLIENTMETRICS structure that receives the information. 
            ' ''' Set the cbSize member of this structure and the <paramref name="uiParam"/> parameter to sizeof(NONCLIENTMETRICS).
            ' ''' </summary>
            ' GetNonClientMetrics = &H29

            ' ''' <summary>
            ' ''' Sets the metrics associated with the nonclient area of nonminimized windows. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a NONCLIENTMETRICS structure that contains the new parameters. 
            ' ''' Set the cbSize member of this structure and the <paramref name="uiParam"/> parameter to sizeof(NONCLIENTMETRICS). 
            ' ''' Also, the lfHeight member of the LOGFONT structure must be a negative value.
            ' ''' </summary>
            ' SetNonClientMetrics = &H2A

            ' ''' <summary>
            ' ''' Retrieves the metrics associated with minimized windows. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a MINIMIZEDMETRICS structure that receives the information. 
            ' ''' Set the cbSize member of this structure and the <paramref name="uiParam"/> parameter to sizeof(MINIMIZEDMETRICS).
            ' ''' </summary>
            ' GetMinimizedMetrics = &H2B

            ' ''' <summary>
            ' ''' Sets the metrics associated with minimized windows. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a MINIMIZEDMETRICS structure that contains the new parameters.
            ' ''' Set the cbSize member of this structure and the <paramref name="uiParam"/> parameter to sizeof(MINIMIZEDMETRICS).
            ' ''' </summary>
            ' SetMinimizedMetrics = &H2C

            ' ''' <summary>
            ' ''' Retrieves the metrics associated with icons. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an ICONMETRICS structure that receives the information. 
            ' ''' Set the cbSize member of this structure and the <paramref name="uiParam"/> parameter to sizeof(ICONMETRICS).
            ' ''' </summary>
            ' GetIconMetrics = &H2D

            ' ''' <summary>
            ' ''' Sets the metrics associated with icons. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an ICONMETRICS structure that contains the new parameters. 
            ' ''' Set the cbSize member of this structure and the <paramref name="uiParam"/> parameter to sizeof(ICONMETRICS).
            ' ''' </summary>
            ' SetIconMetrics = &H2E

            ' ''' <summary>
            ' ''' Retrieves the size of the work area on the primary display monitor. 
            ' ''' The work area is the portion of the screen not obscured by the system taskbar or by application desktop toolbars.
            ' ''' The <paramref name="pvParam"/> parameter must point to a RECT structure that receives
            ' ''' the coordinates of the work area, expressed in virtual screen coordinates.
            ' ''' To get the work area of a monitor other than the primary display monitor, call the GetMonitorInfo function.
            ' ''' </summary>
            ' GetWorkArea = &H30

            ' ''' <summary>
            ' ''' Sets the size of the work area. 
            ' ''' The work area is the portion of the screen not obscured by the system taskbar or by application desktop toolbars.
            ' ''' The <paramref name="pvParam"/> parameter is a pointer to a RECT structure that specifies the new work area rectangle,
            ' ''' expressed in virtual screen coordinates. In a system with multiple display monitors, 
            ' ''' the function sets the work area of the monitor that contains the specified rectangle.
            ' ''' </summary>
            ' SetWorkArea = &H2F

            ' ''' <summary>
            ' ''' This flag is obsolete. 
            ' ''' Previous versions of the system use this flag to determine whether ALT+TAB fast task switching is enabled.
            ' ''' For Windows 95, Windows 98, and Windows NT version 4.0 and later, fast task switching is always enabled.
            ' ''' </summary>
            ' GetFastTaskSwitch = &H23

            ' ''' <summary>
            ' ''' This flag is obsolete. 
            ' ''' Previous versions of the system use this flag to enable or disable ALT+TAB fast task switching.
            ' ''' For Windows 95, Windows 98, and Windows NT version 4.0 and later, fast task switching is always enabled.
            ' ''' </summary>
            ' SetFastTaskSwitch = &H24

            ' ''' <summary>
            ' ''' Retrieves the logical font information for the current icon-title font. 
            ' ''' The <paramref name="uiParam"/> parameter specifies the size of a LOGFONT structure,
            ' ''' and the <paramref name="pvParam"/> parameter must point to the LOGFONT structure to fill in.
            ' ''' </summary>
            ' GetIconTitleLogFont = &H1F

            ' ''' <summary>
            ' ''' Sets the font that is used for icon titles. 
            ' ''' The <paramref name="uiParam"/> parameter specifies the size of a LOGFONT structure,
            ' ''' and the <paramref name="pvParam"/> parameter must point to a LOGFONT structure.
            ' ''' </summary>
            ' SetIconTitleLogFont = &H22

            ' ''' <summary>
            ' ''' Retrieves the current granularity value of the desktop sizing grid. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives the granularity.
            ' ''' </summary>
            ' GetGridGranularity = &H12

            ' ''' <summary>
            ' ''' Sets the granularity of the desktop sizing grid to the value of the <paramref name="uiParam"/> parameter.
            ' ''' </summary>
            ' SetGridGranularity = &H13

            ' ''' <summary>
            ' ''' Sets the current desktop pattern by causing Windows to read the Pattern= setting from the WIN.INI file.
            ' ''' </summary>
            ' SetDeskPattern = &H15

            ' ''' <summary>
            ' ''' Retrieves the two mouse threshold values and the mouse speed.
            ' ''' </summary>
            ' GetMouse = &H3

            ' ''' <summary>
            ' ''' Sets the two mouse threshold values and the mouse speed.
            ' ''' </summary>
            ' SetMouse = &H4

            ' ''' <summary>
            ' ''' Not implemented.
            ' ''' </summary>
            ' LangDriver = &HC

            ' ''' <summary>
            ' ''' Determines whether screen saving is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that 
            ' ''' receives <see langword="True"/> if screen saving is enabled, or <see langword="False"/> otherwise.
            ' ''' Does not work for Windows 7: http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947(v=vs.85).aspx
            ' ''' </summary>
            ' GetScreensaveActive = &H10

            ' ''' <summary>
            ' ''' Retrieves a value that determines whether Windows 8 is displaying apps using the default scaling plateau for the hardware or 
            ' ''' going to the next higher plateau. 
            ' ''' This value is based on the current "Make everything on your screen bigger" setting, 
            ' ''' found in the Ease of Access section of PC settings: 1 is on, 0 is off.
            ' ''' </summary>
            ' GetLogicalDPIoverride = &H9E

            ' ''' <summary>
            ' ''' Do not use.
            ' ''' </summary>
            ' SetLogicalDPIoverride = &H9F

            ' ''' <summary>
            ' ''' Note  When the SPI_SETDESKWALLPAPER flag is used, SystemParametersInfo returns TRUE unless there is an error 
            ' ''' (like when the specified file doesn't exist).
            ' ''' </summary>
            ' SetDeskWallpaper = &H14

            ' ''' <summary>
            ' ''' Retrieves the current contact visualization setting. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a ULONG variable that receives the setting. 
            ' ''' For more information, see Contact Visualization.
            ' ''' </summary>
            ' GetContactVisualization = &H2018

            ' ''' <summary>
            ' ''' Sets the current contact visualization setting. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a ULONG variable that identifies the setting. 
            ' ''' For more information, see Contact Visualization.
            ' ''' </summary>
            ' SetContactVisualization = &H2019

            ' ''' <summary>
            ' ''' Retrieves the current gesture visualization setting. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a ULONG variable that receives the setting. 
            ' ''' For more information, see Gesture Visualization.
            ' ''' </summary>
            ' GetGestureVisualization = &H201A

            ' ''' <summary>
            ' ''' Sets the current gesture visualization setting. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a ULONG variable that identifies the setting. 
            ' ''' For more information, see Gesture Visualization.
            ' ''' </summary>
            ' SetGestureVisualization = &H201B

            ' ''' <summary>
            ' ''' Retrieves the routing setting for wheel button input.
            ' ''' The routing setting determines whether wheel button input is sent to the app with focus (foreground) or the app under the mouse cursor.
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that receives the routing option. 
            ' ''' If the value is zero or MOUSEWHEEL_ROUTING_FOCUS, mouse wheel input is delivered to the app with focus. 
            ' ''' If the value is 1 or MOUSEWHEEL_ROUTING_HYBRID (default), 
            ' ''' mouse wheel input is delivered to the app with focus (desktop apps) or the app under the mouse cursor (Windows Store apps). 
            ' ''' The <paramref name="uiParam"/> parameter is not used.
            ' ''' </summary>
            ' GetMousewheelRouting = &H201C

            ' ''' <summary>
            ' ''' Sets the routing setting for wheel button input. 
            ' ''' The routing setting determines whether wheel button input is sent to the app with focus (foreground) or the app under the mouse cursor.
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that receives the routing option. 
            ' ''' If the value is zero or MOUSEWHEEL_ROUTING_FOCUS, mouse wheel input is delivered to the app with focus. 
            ' ''' If the value is 1 or MOUSEWHEEL_ROUTING_HYBRID (default), 
            ' ''' mouse wheel input is delivered to the app with focus (desktop apps) or the app under the mouse cursor (Windows Store apps). 
            ' ''' Set the <paramref name="uiParam"/> parameter to zero
            ' ''' </summary>
            ' SetMousewheelRouting = &H201D

            ' ''' <summary>
            ' ''' Retrieves the current pen gesture visualization setting. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a ULONG variable that receives the setting. 
            ' ''' For more information, see Pen Visualization.
            ' ''' </summary>
            ' GetPenVisualization = &H201E

            ' ''' <summary>
            ' ''' Sets the current pen gesture visualization setting. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a ULONG variable that identifies the setting. 
            ' ''' For more information, see Pen Visualization.
            ' ''' </summary>
            ' SetPenVisualization = &H201F

            ' ''' <summary>
            ' ''' Starting with Windows 8: Determines whether the active input settings have Local (per-thread, <see langword="True"/>) or Global (session, <see langword="False"/>) scope. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable.
            ' ''' </summary>
            ' GetThreadLocalInputSettings = &H104E

            ' ''' <summary>
            ' ''' Starting with Windows 8: Determines whether the active input settings have Local (per-thread, <see langword="True"/>) or Global (session, <see langword="False"/>) scope. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable, casted by PVOID.
            ' ''' </summary>
            ' SetThreadLocalInputSettings = &H104F

            ' ''' <summary>
            ' ''' Determines whether a window is docked when it is moved to the top, left, or right edges of a monitor or monitor array. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, or <see langword="False"/> otherwise.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetDockMoving = &H90

            ' ''' <summary>
            ' ''' Sets whether a window is docked when it is moved to the top, left, or right docking targets on a monitor or monitor array. 
            ' ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetDockMoving = &H91

            ' ''' <summary>
            ' ''' Determines whether a maximized window is restored when its caption bar is dragged. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, or <see langword="False"/> otherwise.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetDragFromMaximize = &H8C

            ' ''' <summary>
            ' ''' Sets whether a maximized window is restored when its caption bar is dragged. 
            ' ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ' ''' </summary>
            ' SetDragFromMaximize = &H8D

            ' ''' <summary>
            ' ''' Retrieves the threshold in pixels where docking behavior is triggered by using a mouse to drag a window to the edge of a monitor or monitor array. 
            ' ''' The default threshold is 1. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that receives the value.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetMouseDockThreshold = &H7E

            ' ''' <summary>
            ' ''' Sets the threshold in pixels where docking behavior is triggered by using a mouse to drag a window to the edge of a monitor or monitor array. 
            ' ''' The default threshold is 1. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that contains the new value
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetMouseDockThreshold = &H7F

            ' ''' <summary>
            ' ''' Retrieves the threshold in pixels where undocking behavior is triggered by using a mouse to drag a window from the edge of a monitor or 
            ' ''' a monitor array toward the center. 
            ' ''' The default threshold is 20.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetMouseDragoutThreshold = &H84

            ' ''' <summary>
            ' ''' Sets the threshold in pixels where undocking behavior is triggered by using a mouse to drag a window from the edge of a monitor or 
            ' ''' monitor array to its center. 
            ' ''' The default threshold is 20. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that contains the new value.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetMouseDragoutThreshold = &H85

            ' ''' <summary>
            ' ''' Retrieves the threshold in pixels from the top of a monitor or a monitor array where a vertically maximized window is restored when dragged with the mouse. 
            ' ''' The default threshold is 50.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetMouseSideMoveThreshold = &H88

            ' ''' <summary>
            ' ''' Sets the threshold in pixels from the top of the monitor where a vertically maximized window is restored when dragged with the mouse. 
            ' ''' The default threshold is 50. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that contains the new value
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetMouseSideMoveThreshold = &H89

            ' ''' <summary>
            ' ''' Determines whether a window is vertically maximized when it is sized to the top or bottom of a monitor or monitor array. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, or <see langword="False"/> otherwise.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetSnapSizing = &H8E

            ' ''' <summary>
            ' ''' Sets whether a window is vertically maximized when it is sized to the top or bottom of the monitor. 
            ' ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ' ''' Use <see cref="EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetSnapSizing = &H8F

            ' ''' <summary>
            ' ''' Determines whether window arrangement is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, or <see langword="False"/> otherwise.
            ' ''' Window arrangement reduces the number of mouse, pen, or touch interactions needed to move and size top-level windows by 
            ' ''' simplifying the default behavior of a window when it is dragged or sized.
            ' ''' </summary>
            ' GetWinArranging = &H82

            ' ''' <summary>
            ' ''' Sets whether window arrangement is enabled. 
            ' ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ' ''' Window arrangement reduces the number of mouse, pen, or touch interactions needed to move and size top-level windows by 
            ' ''' simplifying the default behavior of a window when it is dragged or sized.
            ' ''' </summary>
            ' SetWinArranging = &H83

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="fWinIni"/> parameter of <see cref="EnvironmentUtil.NativeMethods.SystemParametersInfo"/> function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947(v=vs.85).aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <Flags>
        Friend Enum SystemParametersWinIniFlags As UInteger

            ''' <summary>
            ''' None.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' Writes the new system-wide parameter setting to the user profile.
            ''' </summary>
            UpdateIniFile = &H1

            ''' <summary>
            ''' Broadcasts the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message after updating the user profile.
            ''' </summary>
            SendChange = &H2

            ''' <summary>
            ''' Same as <see cref="EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendChange"/>.
            ''' </summary>
            SendWinIniChange = &H3

        End Enum

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
            ''' The message is sent to all top-level windows in the system, including disabled or invisible unowned windows. 
            ''' The function does not return until each window has timed out. 
            ''' Therefore, the total wait time can be up to the value of uTimeout multiplied by the number of top-level windows.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms644952%28v=vs.85%29.aspx"/>
        ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            HWND_BROADCAST = &HFFFF&

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' A message that is sent to all top-level windows when 
            ''' the SystemParametersInfo function changes a system-wide setting or when policy settings have changed.
            ''' 
            ''' Applications should send <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> to all top-level windows when 
            ''' they make changes to system parameters
            ''' (This message cannot be sent directly to a single window.)
            ''' 
            ''' To send the <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message to all top-level windows, 
            ''' use the <see cref="EnvironmentUtil.NativeMethods.SendMessageTimeout"/> function with the <paramref name="hwnd"/> parameter set to 
            ''' <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.HWND_BROADCAST"/>.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms725497%28v=vs.85%29.aspx"/>
        ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            WM_SETTINGCHANGE = &H1A

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
            WM_COMMAND = &H111UI

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
            ''' Used with <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_COMMAND"/> message.
            ''' </summary>
            MIN_ALL = 419UI

            ''' <summary>
            ''' Undo the minimization of all minimized windows.
            ''' Used with <see cref="EnvironmentUtil.NativeMethods.WindowsMessages.WM_COMMAND"/> message.
            ''' </summary>
            MIN_ALL_UNDO = 416UI

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

#Region " Enumerations "

    ''' <summary>
    ''' Specified an environment scope (registry root key).
    ''' </summary>
    Public Enum EnvironmentScope As Integer

        ''' <summary>
        ''' This reffers to the commonly known "HKLM" or "HKEY_LOCAL_MACHINE" registry root key.
        ''' Changes made on this registry root key will affect all users.
        ''' </summary>
        Machine = 0

        ''' <summary>
        ''' Current User, this reffers to the commonly known "HKCU" or "HKEY_CURRENT_USER" registry root key.
        ''' Changes made on this registry root key will affect the current user.
        ''' </summary>
        CurrentUser = 1

    End Enum

#End Region

#Region " Constructors "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Prevents a default instance of the <see cref="EnvironmentUtil"/> class from being created.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub New()
    End Sub

#End Region

#Region " Child Classes "

#Region " Shell "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related Windows desktop utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class Shell

#Region " Child Classes "

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
            ''' Prevents a default instance of the <see cref="Desktop"/> class from being created.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
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

                EnvironmentUtil.NativeMethods.SendMessage(
                                              EnvironmentUtil.NativeMethods.FindWindow(EnvironmentUtil.Shell.TaskBar.ClassName, String.Empty),
                                              EnvironmentUtil.NativeMethods.WindowsMessages.WM_COMMAND,
                                              New IntPtr(EnvironmentUtil.NativeMethods.WParams.MIN_ALL_UNDO),
                                              New IntPtr(EnvironmentUtil.NativeMethods.LParams.None))

                ' Dim shell As New Shell
                ' shell.UndoMinimizeALL()
                ' shell = Nothing

            End Sub

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Hides the desktop.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
            Public Shared Sub Hide()

                EnvironmentUtil.NativeMethods.SendMessage(
                                              EnvironmentUtil.NativeMethods.FindWindow(EnvironmentUtil.Shell.TaskBar.ClassName, String.Empty),
                                              EnvironmentUtil.NativeMethods.WindowsMessages.WM_COMMAND,
                                              New IntPtr(NativeMethods.WParams.MIN_ALL),
                                              New IntPtr(NativeMethods.LParams.None))

                ' Dim shell As New Shell
                ' shell.MinimizeAll()
                ' shell = Nothing

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
                    shell = Nothing

                Else
                    Throw New NotImplementedException(message:="This feature is not supported in Windows XP.")

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
                shell = Nothing

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
                shell = Nothing

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
                shell = Nothing

            End Sub

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
                    Return New ReadOnlyCollection(Of ShellBrowserWindow)(EnvironmentUtil.Shell.Explorer.GetExplorerWindows.ToList)
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
                    Return New ReadOnlyCollection(Of Folder2)(EnvironmentUtil.Shell.Explorer.GetExplorerWindowsFolders.ToList)
                End Get
            End Property

#End Region

#Region " Constructors "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Prevents a default instance of the <see cref="Explorer"/> class from being created.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
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
                End If

                Dim shell As New Shell32.Shell
                shell.AddToRecent(filePath)
                shell = Nothing

            End Sub

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Refreshes the opened windows explorer folder instances.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
            Public Shared Sub RefreshWindows()

                For Each window As ShellBrowserWindow In EnvironmentUtil.Shell.Explorer.GetExplorerWindows()
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

                shell = Nothing

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

                shell = Nothing

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
            ''' Prevents a default instance of the <see cref="StartMenu"/> class from being created.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
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

            End Sub

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
                    Return EnvironmentUtil.NativeMethods.FindWindow(EnvironmentUtil.Shell.TaskBar.ClassName, Nothing)
                End Get
            End Property

#End Region

#Region " Enumerations "

            ''' <summary>
            ''' Specifies a desktop taskbar visibility flag.
            ''' </summary>
            Private Enum TaskBarVisibility As Integer

                ''' <summary>
                ''' Hides the TaskBar.
                ''' </summary>
                Hide = &H0

                ''' <summary>
                ''' Shows the TaskBar.
                ''' </summary>
                Show = &H5

            End Enum

#End Region

#Region " Constructors "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Prevents a default instance of the <see cref="TaskBar"/> class from being created.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
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

                If (EnvironmentUtil.Shell.TaskBar.SetVisibility(EnvironmentUtil.Shell.TaskBar.TaskBarVisibility.Hide) = 0) AndAlso
                    Not ignoreErrors Then

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

                If (EnvironmentUtil.Shell.TaskBar.SetVisibility(EnvironmentUtil.Shell.TaskBar.TaskBarVisibility.Show) <> 0) AndAlso
                    Not ignoreErrors Then

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
            Private Shared Function SetVisibility(ByVal visibility As EnvironmentUtil.Shell.TaskBar.TaskBarVisibility) As Integer

                Return EnvironmentUtil.NativeMethods.ShowWindow(EnvironmentUtil.Shell.TaskBar.Hwnd, visibility)

            End Function

#End Region

        End Class

#End Region

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Prevents a default instance of the <see cref="EnvironmentVariables"/> class from being created.
        ''' </summary>
        Private Sub New()
        End Sub

#End Region

    End Class

#End Region

#Region " Environment Variables "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related Windows environment variables utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <RegistryPermission(SecurityAction.Demand, Unrestricted:=True)>
    Public NotInheritable Class EnvironmentVariables

#Region " Types "

#Region " EnvironmentVariableInfo "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Defines the info of a Windows environment Variable.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <Serializable>
        Public NotInheritable Class EnvironmentVariableInfo

#Region " Properties "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets or sets the variable name.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The variable name.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public Property Name As String

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets or sets the variable value.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The variable value.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public Property Value As String

#End Region

#Region " Constructors "

            ''' <summary>
            ''' Initializes a new instance of the <see cref="EnvironmentVariableInfo"/> class.
            ''' </summary>
            Public Sub New()
            End Sub

#End Region

        End Class

#End Region

#End Region

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="IEnumerable(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)"/> collection with the environment variables of the specified environment user.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' For Each envVar As EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo In EnvironmentUtil.EnvironmentVariables.CurrentVariables(EnvironmentUtil.EnvironmentScope.CurrentUser)
        ''' 
        '''     Console.WriteLine(String.Format("Name:{0}; Value:{1}", envVar.Name, envVar.Value))
        ''' 
        ''' Next envVar
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A <see cref="IEnumerable(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)"/> collection with the environment variables.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property CurrentVariables(ByVal environmentScope As EnvironmentUtil.EnvironmentScope) As ReadOnlyCollection(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)
            <DebuggerStepThrough>
            Get
                Return New ReadOnlyCollection(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)(EnvironmentUtil.EnvironmentVariables.GetEnvironmentVariables(environmentScope).ToList)
            End Get
        End Property

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Prevents a default instance of the <see cref="EnvironmentVariables"/> class from being created.
        ''' </summary>
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Registers a Windows environment variable.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' EnvironmentUtil.EnvironmentVariables.RegisterVariable(EnvironmentUtil.EnvironmentScope.CurrentUser, "VariableName", "Elektro is the best!", throwOnExistingVariable:=True)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="environmentScope">
        ''' The environment scope that will owns the variable.
        ''' </param>
        ''' 
        ''' <param name="name">
        ''' The variable name.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The variable value.
        ''' </param>
        ''' 
        ''' <param name="throwOnExistingVariable">
        ''' If <see langword="True"/>, raises an exception if the variable already exists.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' name
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Invalid enumeration value;environmentScope
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' The specified variable already exists.;name
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RegisterVariable(ByVal environmentScope As EnvironmentUtil.EnvironmentScope,
                                           ByVal name As String,
                                           ByVal value As String,
                                           Optional ByVal throwOnExistingVariable As Boolean = False)

            If String.IsNullOrWhiteSpace(name) Then
                Throw New ArgumentNullException(paramName:="name")

            Else
                Dim regKey As RegistryKey
                Dim regPath As String

                Select Case environmentScope

                    Case EnvironmentUtil.EnvironmentScope.Machine
                        regKey = Registry.LocalMachine
                        regPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"

                    Case EnvironmentUtil.EnvironmentScope.CurrentUser
                        regKey = Registry.CurrentUser
                        regPath = "Environment\"

                    Case Else
                        Throw New ArgumentException(message:="Invalid enumeration value.", paramName:="environmentScope")

                End Select

                Using regKey

                    If (throwOnExistingVariable) AndAlso
                       (regKey.OpenSubKey(regPath, writable:=False).GetValueNames.
                                   Any(Function(varName As String) varName.ToLower.Equals(name, StringComparison.OrdinalIgnoreCase))) Then

                        Throw New ArgumentException(message:="The specified variable already exists.", paramName:="name")

                    Else
                        regKey.OpenSubKey(regPath, writable:=True).SetValue(name, value, RegistryValueKind.String)
                        EnvironmentUtil.OS.NotifyRegistryChange("Environment")

                    End If

                End Using

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Registers a Windows environment variable.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' EnvironmentUtil.EnvironmentVariables.RegisterVariable(EnvironmentUtil.EnvironmentScope.CurrentUser,
        '''                                                        New EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo With {.Name = "VariableName", .Value = "Elektro is the best!"},
        '''                                                        throwOnExistingVariable:=True)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="environmentScope">
        ''' The environment scope that will owns the variable.
        ''' </param>
        ''' 
        ''' <param name="variableInfo">
        ''' A <see cref="EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo"/> object that contains the variable data.
        ''' </param>
        ''' 
        ''' <param name="throwOnExistingVariable">
        ''' If <see langword="True"/>, raises an exception if the variable already exists.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RegisterVariable(ByVal environmentScope As EnvironmentUtil.EnvironmentScope,
                                           ByVal variableInfo As EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo,
                                           Optional ByVal throwOnExistingVariable As Boolean = False)

            EnvironmentUtil.EnvironmentVariables.RegisterVariable(environmentScope, variableInfo.Name, variableInfo.Value, throwOnExistingVariable)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Unregisters a Windows environment variable.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' EnvironmentUtil.EnvironmentVariables.UnregisterVariable(EnvironmentUtil.EnvironmentScope.CurrentUser, "VariableName", throwOnMissingVariable:=True)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="environmentScope">
        ''' The environment scope that owns the variable.
        ''' </param>
        ''' 
        ''' <param name="name">
        ''' The variable name.
        ''' </param>
        ''' 
        ''' <param name="throwOnMissingVariable">
        ''' If <see langword="True"/>, raises an exception if the variable doesn't exists.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' name
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Invalid enumeration value;environmentScope
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' The specified variable doesn't exists.;name
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub UnregisterVariable(ByVal environmentScope As EnvironmentUtil.EnvironmentScope,
                                             ByVal name As String,
                                             Optional throwOnMissingVariable As Boolean = False)

            If String.IsNullOrWhiteSpace(name) Then
                Throw New ArgumentNullException(paramName:="name")

            Else
                Dim regKey As RegistryKey
                Dim regPath As String

                Select Case EnvironmentScope

                    Case EnvironmentUtil.EnvironmentScope.Machine
                        regKey = Registry.LocalMachine
                        regPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"

                    Case EnvironmentUtil.EnvironmentScope.CurrentUser
                        regKey = Registry.CurrentUser
                        regPath = "Environment\"

                    Case Else
                        Throw New ArgumentException(message:="Invalid enumeration value.", paramName:="environmentScope")

                End Select

                Using regKey

                    If (throwOnMissingVariable) AndAlso
                       Not (regKey.OpenSubKey(regPath, writable:=False).GetValueNames.
                                   Any(Function(varName As String) varName.ToLower.Equals(name, StringComparison.OrdinalIgnoreCase))) Then

                        Throw New ArgumentException(message:="The specified variable doesn't exists.", paramName:="name")

                    Else
                        regKey.OpenSubKey(regPath, writable:=True).DeleteValue(name, throwOnMissingVariable)
                        EnvironmentUtil.OS.NotifyRegistryChange("Environment")

                    End If

                End Using

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Finds an environment variable and returns a <see cref="EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo"/> object that contains the variable data.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exanple>
        ''' Dim envVarInfo As EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo =
        '''     EnvironmentUtil.EnvironmentVariables.GetVariableInfo(EnvironmentUtil.EnvironmentScope.CurrentUser, "System32", throwOnMissingVariable:=False)
        ''' </exanple>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="environmentScope">
        ''' The environment scope that owns the variable.
        ''' </param>
        ''' 
        ''' <param name="name">
        ''' The variable name.
        ''' </param>
        ''' 
        ''' <param name="throwOnMissingVariable">
        ''' If <see langword="True"/>, raises an exception if the variable doesn't exists.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' name
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Invalid enumeration value;environmentScope
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' The specified variable doesn't exists.;name
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo"/> object that contains the variable data.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function GetVariableInfo(ByVal environmentScope As EnvironmentUtil.EnvironmentScope,
                                               ByVal name As String,
                                               Optional throwOnMissingVariable As Boolean = False) As EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo

            If String.IsNullOrWhiteSpace(name) Then
                Throw New ArgumentNullException(paramName:="name")

            Else
                Dim regKey As RegistryKey
                Dim regPath As String

                Select Case environmentScope

                    Case EnvironmentUtil.EnvironmentScope.Machine
                        regKey = Registry.LocalMachine
                        regPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"

                    Case EnvironmentUtil.EnvironmentScope.CurrentUser
                        regKey = Registry.CurrentUser
                        regPath = "Environment\"

                    Case Else
                        Throw New ArgumentException(message:="Invalid enumeration value.", paramName:="environmentScope")

                End Select

                Using regKey

                    If (throwOnMissingVariable) AndAlso
                       Not (regKey.OpenSubKey(regPath, writable:=False).GetValueNames.
                                   Any(Function(varName As String) varName.ToLower.Equals(name, StringComparison.OrdinalIgnoreCase))) Then

                        Throw New ArgumentException(message:="The specified variable doesn't exists.", paramName:="name")

                    Else
                        regKey = regKey.OpenSubKey(regPath, writable:=False)

                        Return (From valueName As String In regKey.GetValueNames
                                Where valueName.Equals(name, StringComparison.OrdinalIgnoreCase)
                                Select New EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo With
                                              {
                                                  .Name = valueName,
                                                  .Value = CStr(regKey.GetValue(valueName, "", RegistryValueOptions.None))
                                              }).FirstOrDefault

                    End If

                End Using

            End If

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns the value of the specified environment variable.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' EnvironmentUtil.EnvironmentVariables.GetValue(EnvironmentUtil.EnvironmentScope.CurrentUser, "System32", throwOnMissingVariable:=True)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="environmentScope">
        ''' The environment scope that owns the variable.
        ''' </param>
        ''' 
        ''' <param name="name">
        ''' The variable name.
        ''' </param>
        ''' 
        ''' <param name="throwOnMissingVariable">
        ''' If <see langword="True"/>, raises an exception if the variable doesn't exists.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' name
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Invalid enumeration value;environmentScope
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' The specified variable doesn't exists.;name
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The value of the specified environment variable.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function GetValue(ByVal environmentScope As EnvironmentUtil.EnvironmentScope,
                                        ByVal name As String,
                                        Optional throwOnMissingVariable As Boolean = False) As String

            If String.IsNullOrWhiteSpace(name) Then
                Throw New ArgumentNullException(paramName:="name")

            Else
                Dim regKey As RegistryKey
                Dim regPath As String

                Select Case environmentScope

                    Case EnvironmentUtil.EnvironmentScope.Machine
                        regKey = Registry.LocalMachine
                        regPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"

                    Case EnvironmentUtil.EnvironmentScope.CurrentUser
                        regKey = Registry.CurrentUser
                        regPath = "Environment\"

                    Case Else
                        Throw New ArgumentException(message:="Invalid enumeration value.", paramName:="environmentScope")

                End Select

                Using regKey

                    If (throwOnMissingVariable) AndAlso
                       Not (regKey.OpenSubKey(regPath, writable:=False).GetValueNames.
                                   Any(Function(varName As String) varName.ToLower.Equals(name, StringComparison.OrdinalIgnoreCase))) Then

                        Throw New ArgumentException(message:="The specified variable doesn't exists.", paramName:="name")

                    Else
                        Return CStr(regKey.OpenSubKey(regPath, writable:=False).GetValue(name, "", RegistryValueOptions.None))

                    End If

                End Using

            End If

        End Function

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="IEnumerable(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)"/> collection with the 
        ''' environment variables of the specified environment user.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="environmentScope">
        ''' The environment scope that owns the variable.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentException">
        ''' Invalid enumeration value;environmentScope
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="IEnumerable(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)"/> collection with the environment variables.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Shared Function GetEnvironmentVariables(ByVal environmentScope As EnvironmentUtil.EnvironmentScope) As IEnumerable(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)

            Dim regKey As RegistryKey
            Dim regPath As String

            Select Case environmentScope

                Case EnvironmentUtil.EnvironmentScope.Machine
                    regKey = Registry.LocalMachine
                    regPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"

                Case EnvironmentUtil.EnvironmentScope.CurrentUser
                    regKey = Registry.CurrentUser
                    regPath = "Environment\"

                Case Else
                    Throw New ArgumentException(message:="Invalid enumeration value.", paramName:="environmentScope")

            End Select

            Using regKey

                regKey = regKey.OpenSubKey(regPath, writable:=False)

                Return From valueName As String In regKey.GetValueNames
                       Select New EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo With
                                      {
                                          .Name = valueName,
                                          .Value = CStr(regKey.GetValue(valueName, "", RegistryValueOptions.None))
                                      }

            End Using

        End Function

#End Region

    End Class

#End Region

#Region " FileSystem "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related windows file system utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class FileSystem

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="FileSystem"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether a directory name or file name contains invalid windows path characters.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemName">
        ''' The item name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if item contains invalid windows name characters, <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ItemNameIsInvalid(ByVal itemName As String) As Boolean

            Return itemName.Any(Function(c As Char) Path.GetInvalidFileNameChars.Contains(c))

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether a directory path or a file path contains invalid windows path characters.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemPath">
        ''' The item path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if item contains invalid windows path characters, <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ItemPathIsInvalid(ByVal itemPath As String) As Boolean

            Return itemPath.Any(Function(c As Char) Path.GetInvalidPathChars.Contains(c))

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified item is a name or a path, 
        ''' then, determines whether the item contains invalid windows name or path characters.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemNameOrPath">
        ''' The item name or path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' <see langword="True"/> if item contains invalid windows name characters, <see langword="False"/> otherwise.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function ItemNameOrPathIsInvalid(ByVal itemNameOrPath As String) As Boolean

            Try
                Path.GetDirectoryName(itemNameOrPath)
                ' It's a item path.
                Return EnvironmentUtil.FileSystem.ItemPathIsInvalid(itemNameOrPath)

            Catch ex As ArgumentException
                ' It's a item name.
                Return EnvironmentUtil.FileSystem.ItemNameIsInvalid(itemNameOrPath)

            End Try

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the avaliable item verbs of the specified file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemPath">
        ''' The item path.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' An <see cref="IEnumerable(Of FolderItemVerb)"/> containing the avaliable item verbs.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Function GetItemVerbs(ByVal itemPath As String) As IEnumerable(Of FolderItemVerb)

            Dim shell As New Shell32.Shell
            Dim link As FolderItem = shell.NameSpace(Path.GetDirectoryName(itemPath)).ParseName(Path.GetFileName(itemPath))

            Return link.Verbs.Cast(Of FolderItemVerb)()

        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Invokes an item verb on the specified file or directory.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="itemPath">
        ''' The item path.
        ''' </param>
        ''' 
        ''' <param name="verbName">
        ''' The verb name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub InvokeItemVerb(ByVal itemPath As String,
                                         ByVal verbName As String)

            Dim shell As New Shell32.Shell
            Dim link As FolderItem = shell.NameSpace(Path.GetDirectoryName(itemPath)).ParseName(Path.GetFileName(itemPath))

            link.InvokeVerb(verbName)

        End Sub

#End Region

    End Class

#End Region

#Region " Operating System "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related operating system utilities.
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
        ''' An <see cref="Architecture"></see> object that specifies the architecture of the current operating system.
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

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the current screensaver filepath.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The current screensaver filepath.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ScreensaverPath As String

            <DebuggerStepThrough>
            Get
                Using regkey As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop")
                    Return CStr(regkey.GetValue("SCRNSAVE.EXE", String.Empty, RegistryValueOptions.None))
                End Using
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As String)
                Using regkey As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop")
                    regkey.SetValue("SCRNSAVE.EXE", value, RegistryValueKind.String)
                End Using
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the number of milliseconds that the service control manager waits before 
        ''' terminating a service that does not respond to a shutdown request.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The number of milliseconds that the service control manager waits before 
        ''' terminating a service that does not respond to a shutdown request.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 2147483647;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property WaitToKillServiceTimeout As Integer

            <DebuggerStepThrough>
            Get
                Dim result As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWaitToKillServiceTimeout,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 2147483647.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetWaitToKillServiceTimeout,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the number of milliseconds that the system waits before 
        ''' terminating an application that does not respond to a shutdown request.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The number of milliseconds that the system waits before 
        ''' terminating an application that does not respond to a shutdown request.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 2147483647;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property WaitToKillAppTimeout As Integer

            <DebuggerStepThrough>
            Get
                Dim result As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWaitToKillTimeout,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    '   Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 2147483647.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetWaitToKillTimeout,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the number of milliseconds that a thread can go without dispatching a message before the system considers it unresponsive.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The number of milliseconds that a thread can go without dispatching a message before the system considers it unresponsive.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 2147483647;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property HungAppTimeout As Integer

            <DebuggerStepThrough>
            Get
                Dim result As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetHungAppTimeout,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 2147483647.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetHungAppTimeout,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether the screen saver requires a password to display the Windows desktop.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the screen saver requires a password to display the Windows desktop, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ScreensaveSecureEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetScreensaveSecure,
                                                                0UI,
                                                                result,
                                                                EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetScreensaveSecure,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the number of characters to scroll when the horizontal mouse wheel is moved.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The number of characters to scroll when the horizontal mouse wheel is moved.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 2147483647;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property WheelscrollChars As Integer

            <DebuggerStepThrough>
            Get
                Dim result As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetWheelscrollChars,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 2147483647.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetWheelscrollChars,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether the system language bar is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if the system language bar is enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property SystemLanguageBarEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetSystemlanguageBar,
                                                                0UI,
                                                                result,
                                                                EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                Dim int32value As Integer = 0
                If value Then
                    int32value = 1
                End If
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetSystemlanguageBar,
                                                                          0UI,
                                                                          int32value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether ClearType feature is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if ClearType is enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property CleartypeEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetCleartype,
                                                                0UI,
                                                                result,
                                                                EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    '  Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                Dim int32value As Integer = 0
                If value Then
                    int32value = 1
                End If
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetCleartype,
                                                                          0UI,
                                                                          int32value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether animation effects in the client area of applications are enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if effects are enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ClientAreaAnimationEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetClientAreaAnimation,
                                                                0UI,
                                                                result,
                                                                EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetClientAreaAnimation,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether overlapped content is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if overlapped content is enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property OverlappedContentEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetDisableOverlappedContent,
                                                                0UI,
                                                                result,
                                                                EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDisableOverlappedContent,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that enables or disables the font smoothing feature, 
        ''' which uses font antialiasing to make font curves appear smoother by painting pixels at different gray levels.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if font smoothing is enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property FontSmoothingEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsFontSmoothingEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetFontSmoothing,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that enables or disables dragging of full windows.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if dragging is enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property DragFullWindowsEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.DragFullWindows
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDragFullWindows,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that swaps or restores the meaning of the left and right mouse buttons.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if mouse swap is enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseButtonsSwapEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.MouseButtonsSwapped
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMousebuttonSwap,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether UI effects are enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if UI effects are enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property UIEffectsEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.UIEffectsEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetUiEffects,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether an application can reset the screensaver's timer by calling the SendInput function
        ''' to simulate keyboard or mouse input.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property BlockSendInputResetsEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetBlockSendInputResets,
                                                                          0I,
                                                                          result,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetBlockSendInputResets,
                                                                          Not value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether flat menu appearance for native User menus are enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property FlatMenuEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsFlatMenuEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetFlatMenu,
                                                                          0I,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether mouse vanish feature is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseVanishEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetMouseVanish,
                                                                          0I,
                                                                          result,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseVanish,
                                                                          0I,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether mouse clicklock feature is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseClickLockEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetMouseClickLock,
                                                                          0I,
                                                                          result,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseClickLock,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether mouse sonar feature is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseSonarEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetMouseSonar,
                                                                          0I,
                                                                          result,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseSonar,
                                                                          0I,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether the cursor has a shadow around it.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property CursorShadowEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetCursorShadow,
                                                                          0I,
                                                                          result,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetCursorShadow,
                                                                          0I,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether drop shadow effect feature is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property DropShadowEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsDropShadowEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDropShadow,
                                                                          0I,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether a selected menu item by the user should remain on the screen briefly while fading out
        ''' after the menu is dismissed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property SelectionFadeEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsSelectionFadeEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetSelectionFade,
                                                                          0I,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether menu fade animation is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MenuFadeEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsMenuFadeEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMenuFade,
                                                                          0I,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether hot tracking of user-interface elements such as menu names on menu bars is enabled.
        ''' Hot-tracking means that when the cursor moves over an item, it is highlighted but not selected.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property HotTrackingEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsHotTrackingEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetHotTracking,
                                                                          0I,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether underlining of menu access key letters is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MenuAccessKeysUnderlined As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.MenuAccessKeysUnderlined
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetKeyboardCues,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether the gradient effect for window title bars are enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property TitleBarGradientEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsTitleBarGradientEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetGradientCaptions,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether the smooth-scrolling effect for list boxes is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ListBoxSmoothScrollingEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsListBoxSmoothScrollingEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetListboxSmoothScrolling,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether ToolTip animations are enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ToolTipAnimationEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsToolTipAnimationEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetTooltipAnimation,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether the slide-open effect for combo boxes is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ComboBoxAnimationEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsComboBoxAnimationEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetComboboxAnimation,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether menu animations are enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MenuAnimationEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsMenuAnimationEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMenuAnimation,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether active window tracking (activating the window the mouse is on) is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ActiveWindowTrackingEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsActiveWindowTrackingEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetActiveWindowTracking,
                                                                          0UI,
                                                                          value,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether the snap-to-default-button feature is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property SnapToDefaultEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsSnapToDefaultEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetSnapToDefButton,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether cursor movements shows a trail of cursors.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property MouseTrailEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetMouseTrails,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If

                Select Case result

                    Case 0UI, 1UI
                        Return False

                    Case Else
                        Return True

                End Select

            End Get

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether screensaver is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ScreensaverEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Using regkey As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop")
                    Return CBool(regkey.GetValue("ScreenSaveActive", False, RegistryValueOptions.None))
                End Using
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetScreensaveActive,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether icon-title wrapping is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property IconTitleWrappingEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IsIconTitleWrappingEnabled
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetIconTitleWrap,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets a value that determines whether system beep is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if enabled, <see langword="False"/> otherwise.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property BeepEnabled As Boolean

            <DebuggerStepThrough>
            Get
                Dim result As Boolean
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetBeep,
                                                                          0UI,
                                                                          result,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If

                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetBeep,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.UpdateIniFile) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the number of times SetForegroundWindow will flash the taskbar button when rejecting a foreground switch request.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The number of times SetForegroundWindow will flash the taskbar button when rejecting a foreground switch request.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 65535.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ForegroundFlashCount As UShort

            <DebuggerStepThrough>
            Get
                Dim result As UShort
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetForegroundFlashCount,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As UShort)
                If (value < 0) OrElse (value > UShort.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 65535.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetForegroundFlashCount,
                                                                              0I,
                                                                              CUInt(value),
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the active window tracking delay, in milliseconds.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The active window tracking delay, in milliseconds.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 65535.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ActiveWindowTrackingTimeout As UShort

            <DebuggerStepThrough>
            Get
                Dim result As UShort
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetActiveWndTrkTimeout,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As UShort)
                If (value < 0) OrElse (value > UShort.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 65535.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetActiveWndTrkTimeout,
                                                                              0I,
                                                                              CUInt(value),
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the amount of time following user input, in milliseconds, 
        ''' during which the system will not allow applications to force themselves into the foreground.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The amount of time following user input, in milliseconds, 
        ''' during which the system will not allow applications to force themselves into the foreground.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 65535.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ForegroundLockTimeout As UShort

            <DebuggerStepThrough>
            Get
                Dim result As UShort
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetForegroundLockTimeout,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As UShort)
                If (value < 0) OrElse (value > UShort.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 65535.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetForegroundLockTimeout,
                                                                              0I,
                                                                              CUInt(value),
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the border multiplier factor that determines the width of a window's sizing border.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The border multiplier factor that determines the width of a window's sizing border.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property BorderMultiplierFactor As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.BorderMultiplierFactor
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetBorder,
                                                                          value,
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the contrast value used in ClearType smoothing. From 1000 to 2200.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The contrast value used in ClearType smoothing. From 1000 to 2200.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 1000 and 2200.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property FontSmoothingContrast As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.FontSmoothingContrast
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 1000I) OrElse (value > 2200I) Then
                    Throw New ArgumentException(message:="Value should be between 1000 and 2200.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetFontSmoothingContrast,
                                                                              0UI,
                                                                              value,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange Or
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.UpdateIniFile) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time delay before the primary mouse button is locked.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The time delay before the primary mouse button is locked.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be grater than -1.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseClickLockTime As Integer

            <DebuggerStepThrough>
            Get
                Dim result As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetMouseClickLockTime,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) Then
                    Throw New ArgumentException(message:="Value should be grater than -1.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseClickLockTime,
                                                                              0UI,
                                                                              value,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the width, in pixels, of the caret in edit controls.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The width, in pixels, of the caret in edit controls.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 100.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property CaretWidth As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.CaretWidth
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > 100) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 100.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetCaretWidth,
                                                                              0UI,
                                                                              value,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the current mouse speed. From 1 to 20.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The current mouse speed. From 1 to 20.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 1 and 20.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseSpeed As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.MouseSpeed
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 1) OrElse (value > 20) Then
                    Throw New ArgumentException(message:="Value should be between 1 and 20.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseSpeed,
                                                                              0UI,
                                                                              value,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time, in milliseconds, that the system waits before displaying a 
        ''' cascaded shortcut menu when the mouse cursor is over a submenu item.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The time, in milliseconds, that the system waits before displaying a 
        ''' cascaded shortcut menu when the mouse cursor is over a submenu item.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be greater than 0.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MenuShowDelay As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.MenuShowDelay
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) Then
                    Throw New ArgumentException(message:="Value should be greater than -1.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMenuShowDelay,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the number of lines to scroll when the mouse wheel is rotated.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The number of lines to scroll when the mouse wheel is rotated.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be greater than 0.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseWheelScrollLines As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.MouseWheelScrollLines
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 1) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be greater than 0.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetWheelScrollLines,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time, in milliseconds, that the mouse pointer has to stay in the hover rectangle for TrackMouseEvent
        ''' to generate a WM_MOUSEHOVER message.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The time, in milliseconds, that the mouse pointer has to stay in the hover rectangle for TrackMouseEvent
        ''' to generate a WM_MOUSEHOVER message.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be greater than 0.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseHoverTime As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.MouseHoverTime
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 1) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be greater than 0.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseHoverTime,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the number of cursors drawn when mouse trail feature is enabled.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The number of cursors drawn when mouse trail feature is enabled.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 16.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseTrailAmount As Integer

            <DebuggerStepThrough>
            Get
                Dim result As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetMouseTrails,
                                                              0UI,
                                                              result,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > 16) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 16.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseTrails,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the maximum number of milliseconds that can occur between the first and second clicks of a double-click.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The maximum number of milliseconds that can occur between the first and second clicks of a double-click.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be greater thaan 0.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property DoubleClickTime As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.DoubleClickTime
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 1) Then
                    Throw New ArgumentException(message:="Value should be greater than 0.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDoubleclickTime,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the keyboard repeat-delay. From 0 to 3.
        ''' Where zero sets the shortest delay (approximately 250 ms) and 3 sets the longest delay (approximately 1 second).
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The keyboard repeat-delay, in seconds. From 0 to 3.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 3.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property KeyboardDelay As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.KeyboardDelay
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > 3) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 3.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetKeyboardDelay,
                                                                              CUInt(value),
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the screen saver time-out, in seconds. From 1 to 599940.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The screen saver time-out, in seconds. From 1 to 599940.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 1 and 599940.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property ScreensaverTimeout As Integer

            <DebuggerStepThrough>
            Get
                Dim result As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetScreensaveTimeout,
                                                                          0UI,
                                                                          result,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 1) OrElse (value > 599940) Then
                    Throw New ArgumentException(message:="Value should be between 1 and 599940.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetScreensaveTimeout,
                                                                              value,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the keyboard repeat-speed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The keyboard repeat-speed, from 0 to 31.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Value should be between 0 and 31.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property KeyboardSpeed As Integer

            <DebuggerStepThrough>
            Get
                Return SystemInformation.KeyboardSpeed
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > 31) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 31.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetKeyboardSpeed,
                                                                              CUInt(value),
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the time that notification pop-ups should be displayed, in seconds.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The time that notification pop-ups should be displayed, in seconds.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MessageDuration As ULong

            <DebuggerStepThrough>
            Get
                Dim result As ULong
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetMessageDuration,
                                                               0UI,
                                                               result,
                                                               EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CULng(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As ULong)
                If (value < 5) OrElse (value > 4294967295) Then
                    Throw New ArgumentException(message:="Value should be between 5 and 4294967295.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMessageDuration,
                                                                              0UI,
                                                                              CLng(value),
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the width and height, in pixels, of the focus rectangle drawn with DrawFocusRect.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The width and height, in pixels, of the focus rectangle drawn with DrawFocusRect.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Width should be greater than 0.;value
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Height should be greater than 0.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property FocusBorderSize As Size

            <DebuggerStepThrough>
            Get
                Dim width As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetFocusBorderWidth,
                                                              0UI,
                                                              width,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If

                Dim height As UInteger
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetFocusBorderHeight,
                                                              0UI,
                                                              height,
                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return New Size(CInt(width), CInt(height))

            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Size)
                If (value.Width < 1) Then
                    Throw New ArgumentException(message:="Width should be greater than 0.", paramName:="value")

                ElseIf (value.Height < 1) Then
                    Throw New ArgumentException(message:="Height should be greater than 0.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetFocusBorderWidth,
                                                                              0UI,
                                                                              value.Width,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetFocusBorderHeight,
                                                                              0UI,
                                                                              value.Height,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the width and height, in pixels, of the rectangle which the mouse pointer has to stay for TrackMouseEvent
        ''' to generate a WM_MOUSEHOVER message.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The width and height, in pixels, of the rectangle which the mouse pointer has to stay for TrackMouseEvent
        ''' to generate a WM_MOUSEHOVER message.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Width should be greater than 0.;value
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Height should be greater than 0.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property MouseHoverSize As Size

            <DebuggerStepThrough>
            Get
                Return SystemInformation.MouseHoverSize
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Size)
                If (value.Width < 1) Then
                    Throw New ArgumentException(message:="Width should be greater than 0", paramName:="value")

                ElseIf (value.Height < 1) Then
                    Throw New ArgumentException(message:="Height should be greater than 0.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseHoverWidth,
                                                                              value.Width,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMouseHoverHeight,
                                                                              value.Height,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the width and height, in pixels, of the rectangle used to detect the start of a drag operation.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The width and height, in pixels, of the rectangle used to detect the start of a drag operation.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Width should be greater than 0.;value
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Height should be greater than 0.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property DragSize As Size

            <DebuggerStepThrough>
            Get
                Return SystemInformation.DragSize
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Size)
                If (value.Width < 1) Then
                    Throw New ArgumentException(message:="Width should be greater than 0.", paramName:="value")

                ElseIf (value.Height < 1) Then
                    Throw New ArgumentException(message:="Height should be greater than 0.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDragWidth,
                                                                              value.Width,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDragHeight,
                                                                              value.Height,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the width and height, in pixels, of the double-click rectangle which 
        ''' the second click of a double-click must fall for it to be registered as a double-click.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The width and height, in pixels, of the double-click rectangle which 
        ''' the second click of a double-click must fall for it to be registered as a double-click.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Width should be greater than 31.;value
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Height should be greater than 31.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property DoubleClickSize As Size

            <DebuggerStepThrough>
            Get
                Return SystemInformation.DoubleClickSize
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Size)
                If (value.Width < 32) Then
                    Throw New ArgumentException(message:="Width should be greater than 31.", paramName:="value")

                ElseIf (value.Height < 32) Then
                    Throw New ArgumentException(message:="Height should be greater than 31.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDoubleClickWidth,
                                                                              value.Width,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDoubleClickHeight,
                                                                              value.Height,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the width and height, in pixels, of an icon cell.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The width and height, in pixels, of an icon cell.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Width should be greater than 31.;value
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Height should be greater than 31.;value
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property IconSpacing As Size

            <DebuggerStepThrough>
            Get
                Return SystemInformation.IconSpacingSize
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Size)
                If (value.Width < 32) Then
                    Throw New ArgumentException(message:="Width should be greater than 31.", paramName:="value")

                ElseIf (value.Height < 32) Then
                    Throw New ArgumentException(message:="Height should be greater than 31.", paramName:="value")

                Else
                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.IconHorizontalSpacing,
                                                                              value.Width,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.IconVerticalSpacing,
                                                                              value.Height,
                                                                              0UI,
                                                                              EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the side of pop-up menus that are aligned to the corresponding menu-bar item.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The side of pop-up menus that are aligned to the corresponding menu-bar item.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property PopupMenuAlignment As LeftRightAlignment

            <DebuggerStepThrough>
            Get
                Return SystemInformation.PopupMenuAlignment
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As LeftRightAlignment)
                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetMenuDropAlignment,
                                                                          CBool(value),
                                                                          0UI,
                                                                          EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
            End Set

        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets or sets the system date and time.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' Dim dateAndTime As New Date(year:=2000, month:=1, day:=1,
        '''                             hour:=0, minute:=0, second:=0)
        ''' 
        ''' Dim dateOnly As New Date(year:=2000, month:=1, day:=1,
        '''                          hour:=Date.Now.Hour, minute:=Date.Now.Minute, second:=Date.Now.Second)
        ''' 
        ''' Dim timeOnly As New Date(year:=Date.Today.Year, month:=Date.Today.Month, day:=Date.Today.Day,
        '''                          hour:=0, minute:=0, second:=0)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The system date and time.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared Property SystemDateTime() As Date
            <DebuggerStepThrough>
            Get
                Return DateTime.Now
            End Get
            <DebuggerStepThrough>
            Set(ByVal value As Date)

                If value.Second = 0 Then
                    value.AddSeconds(1)
                End If

                ' Set the System Hour.
                Microsoft.VisualBasic.TimeOfDay = value

                ' Set the System Date.
                Microsoft.VisualBasic.DateString = value.ToString("MM/dd/yyyy")

            End Set
        End Property

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Indicates the possible processor architectures.
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

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="OS"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system to update after a registry change.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="keyName">
        ''' A string that indicates the area containing the system parameter that was changed.
        ''' 
        ''' This string can be the name of a registry key or the name of a section in the Win.ini file. 
        ''' 
        ''' When the string is a registry name, it typically indicates only the leaf node in the registry, not the full path.
        ''' 
        ''' To effect a change in the policy settings, this parameter points to the string "Policy".
        ''' To effect a change in the locale settings, this parameter points to the string "intl".
        ''' To effect a change in the environment variables for the system or the user, this parameter points to the string "Environment".
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms725497%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyRegistryChange(ByVal keyName As String)

            EnvironmentUtil.NativeMethods.SendMessageTimeout(New IntPtr(EnvironmentUtil.NativeMethods.WindowsMessages.HWND_BROADCAST),
                                                             CInt(EnvironmentUtil.NativeMethods.WindowsMessages.WM_SETTINGCHANGE),
                                                             New IntPtr(0),
                                                             keyName,
                                                             EnvironmentUtil.NativeMethods.SendMessageTimeoutFlags.AbortIfHung,
                                                             1,
                                                             IntPtr.Zero)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reloads the system cursors.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub ReloadSystemCursors()

            If EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.Setcursors,
                                                                      0UI,
                                                                      0UI,
                                                                      EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Reloads the system icons.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub ReloadSystemIcons()

            If EnvironmentUtil.NativeMethods.SystemParametersInfo(EnvironmentUtil.NativeMethods.SystemParametersActionFlags.Seticons,
                                                                      0UI,
                                                                      0UI,
                                                                      EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a directory has been created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="directoryPath">
        ''' The full path of the directory that has been created.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' directoryPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyDirectoryCreated(ByVal directoryPath As String)

            If String.IsNullOrWhiteSpace(directoryPath) Then
                Throw New ArgumentNullException(paramName:="directoryPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.DirectoryCreated,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=directoryPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a directory has been deleted.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="directoryPath">
        ''' The full path of the directory that has been deleted.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' directoryPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyDirectoryDeleted(ByVal directoryPath As String)

            If String.IsNullOrWhiteSpace(directoryPath) Then
                Throw New ArgumentNullException(paramName:="directoryPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.DirectoryDeleted,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=directoryPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a directory has been renamed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="oldDirectoryPath">
        ''' The previous full path of the directory that has been renamed.
        ''' </param>
        ''' 
        ''' <param name="newDirectoryPath">
        ''' The new full path of the directory that has been renamed.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' oldDirectoryPath or newDirectoryPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyDirectoryRenamed(ByVal oldDirectoryPath As String,
                                                 ByVal newDirectoryPath As String)

            If String.IsNullOrWhiteSpace(oldDirectoryPath) Then
                Throw New ArgumentNullException(paramName:="oldDirectoryPath")

            ElseIf String.IsNullOrWhiteSpace(newDirectoryPath) Then
                Throw New ArgumentNullException(paramName:="newDirectoryPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.DirectoryRenamed,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=oldDirectoryPath, dwItem2:=newDirectoryPath)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a file has been created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filePath">
        ''' The full path of the file that has been created.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' filePath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyFileCreated(ByVal filePath As String)

            If String.IsNullOrWhiteSpace(filePath) Then
                Throw New ArgumentNullException(paramName:="filePath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.ItemCreated,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=filePath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a file has been deleted.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filePath">
        ''' The full path of the file that has been deleted.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' filePath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyFileDeleted(ByVal filePath As String)

            If String.IsNullOrWhiteSpace(filePath) Then
                Throw New ArgumentNullException(paramName:="filePath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.ItemDeleted,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=filePath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a file has been renamed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="oldFilePath">
        ''' The previous full path of the file that has been renamed.
        ''' </param>
        ''' 
        ''' <param name="newFilePath">
        ''' The new full path of the file that has been renamed.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' oldFilePath or newFilePath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyFileRenamed(ByVal oldFilePath As String,
                                            ByVal newFilePath As String)

            If String.IsNullOrWhiteSpace(oldFilePath) Then
                Throw New ArgumentNullException(paramName:="oldFilePath")

            ElseIf String.IsNullOrWhiteSpace(newFilePath) Then
                Throw New ArgumentNullException(paramName:="newFilePath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.ItemRenamed,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=oldFilePath, dwItem2:=newFilePath)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a drive has been added.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="driveRootPath">
        ''' The root path of the drive that has been added.
        ''' </param>
        ''' 
        ''' <param name="createShellWindow">
        ''' If <see langword="True"/>, tell the Shell to create a new window for the drive.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' driveRootPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyDriveAdded(ByVal driveRootPath As String,
                                           Optional ByVal createShellWindow As Boolean = False)

            If String.IsNullOrWhiteSpace(driveRootPath) Then
                Throw New ArgumentNullException(paramName:="driveRootPath")

            Else
                Select Case createShellWindow

                    Case True
                        EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.DriveAddedShell,
                                                                     EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                                     dwItem1:=driveRootPath, dwItem2:=Nothing)
                    Case Else
                        EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.DriveAdded,
                                                                     EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                                     dwItem1:=driveRootPath, dwItem2:=Nothing)

                End Select

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a drive has been removed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="driveRootPath">
        ''' The root path of the drive that has been removed.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' driveRootPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyDriveRemoved(ByVal driveRootPath As String)

            If String.IsNullOrWhiteSpace(driveRootPath) Then
                Throw New ArgumentNullException(paramName:="driveRootPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.DriveRemoved,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=driveRootPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a storage media has been inserted into a drive.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="driveRootPath">
        ''' The root path of the drive that contains the new media.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' driveRootPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyMediaInserted(ByVal driveRootPath As String)

            If String.IsNullOrWhiteSpace(driveRootPath) Then
                Throw New ArgumentNullException(paramName:="driveRootPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.MediaInserted,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=driveRootPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a storage media has been removed from a drive.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="driveRootPath">
        ''' The root path of the drive from which the media was removed.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' driveRootPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyMediaRemoved(ByVal driveRootPath As String)

            If String.IsNullOrWhiteSpace(driveRootPath) Then
                Throw New ArgumentNullException(paramName:="driveRootPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.MediaRemoved,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=driveRootPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a directory on the local computer is being shared via the network.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="directoryPath">
        ''' The full path of the directory that is being shared.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' directoryPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyNetworkFolderShared(ByVal directoryPath As String)

            If String.IsNullOrWhiteSpace(directoryPath) Then
                Throw New ArgumentNullException(paramName:="directoryPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.NetShared,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=directoryPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a directory on the local computer is no longer being shared via the network.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="directoryPath">
        ''' The full path of the directory that is being not shared.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' directoryPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyNetworkFolderUnshared(ByVal directoryPath As String)

            If String.IsNullOrWhiteSpace(directoryPath) Then
                Throw New ArgumentNullException(paramName:="directoryPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.NetUnshared,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=directoryPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that the attributes of a file have changed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filePath">
        ''' The full path of the file on which its attributes has chaged.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' filePath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyFileAttributesChanged(ByVal filePath As String)

            If String.IsNullOrWhiteSpace(filePath) Then
                Throw New ArgumentNullException(paramName:="filePath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.ItemAttributesChanged,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=filePath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that the attributes of a directory have changed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="directoryPath">
        ''' The full path of the directory on which its attributes has chaged.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' directoryPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyDirectoryAttributesChanged(ByVal directoryPath As String)

            If String.IsNullOrWhiteSpace(directoryPath) Then
                Throw New ArgumentNullException(paramName:="directoryPath")

            Else
                EnvironmentUtil.OS.NotifyFileAttributesChanged(directoryPath)

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that the contents of an existing folder have changed but the folder still exists and has not been renamed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="directoryPath">
        ''' The full path of the directory that has chaged.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' directoryPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyUpdateDirectory(ByVal directoryPath As String)

            If String.IsNullOrWhiteSpace(directoryPath) Then
                Throw New ArgumentNullException(paramName:="directoryPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.UpdateDirectory,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=directoryPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that a file type association has changed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyFileAssociationChanged()

            EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.FileAssocChanged,
                                                         EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.ItemIdList,
                                                         dwItem1:=Nothing, dwItem2:=Nothing)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that amount of free space on a drive has changed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="driveRootPath">
        ''' The root path of the drive on which the free space changed.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' driveRootPath
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyFreespaceChanged(ByVal driveRootPath As String)

            If String.IsNullOrWhiteSpace(driveRootPath) Then
                Throw New ArgumentNullException(paramName:="driveRootPath")

            Else
                EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.FreespaceChanged,
                                                             EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.PathA,
                                                             dwItem1:=driveRootPath, dwItem2:=Nothing)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system that an image in the system image list has changed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub NotifyUpdateImage()

            EnvironmentUtil.NativeMethods.SHChangeNotify(EnvironmentUtil.NativeMethods.SHChangeNotifyEventID.UpdateImage,
                                                         EnvironmentUtil.NativeMethods.SHChangeNotifyFlags.Dword,
                                                         dwItem1:=Nothing, dwItem2:=Nothing)

        End Sub

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
            shell = Nothing

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
            shell = Nothing

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
            shell = Nothing

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
            shell = Nothing

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
            shell = Nothing

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
            shell = Nothing

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
            shell = Nothing

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
            shell = Nothing

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
            shell = Nothing

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
            shell = Nothing

        End Sub

#End Region

    End Class

#End Region

#Region " Programs "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related Windows programs utilities.
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
        ''' Prevents a default instance of the <see cref="Programs"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

    End Class

#End Region

#Region " Theming "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related theming/personalization utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class Theming

#Region " Types "

#Region " ThemeInfo "

        ''' <summary>
        ''' Defined the information of a Windows Visual Theme.
        ''' </summary>
        <Serializable>
        Public NotInheritable Class ThemeInfo

#Region " Properties "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets the theme filename.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The theme filename.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public ReadOnly Property FileName() As String
                <DebuggerStepThrough>
                Get
                    Return Path.GetFileName(Me.filepathB)
                End Get
            End Property

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets the theme filepath.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The theme filepath.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public ReadOnly Property Filepath() As String
                <DebuggerStepThrough>
                Get
                    Return Me.filepathB
                End Get
            End Property
            ''' <summary>
            ''' ( Backing Field )
            ''' The theme filepath.
            ''' </summary>
            Private ReadOnly filepathB As String

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets the theme color scheme name.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The theme color scheme name.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public ReadOnly Property ColorSchemeName() As String
                <DebuggerStepThrough>
                Get
                    Return Me.colorSchemeNameB
                End Get
            End Property
            ''' <summary>
            ''' ( Backing Field )
            ''' The theme color scheme name.
            ''' </summary>
            Private ReadOnly colorSchemeNameB As String

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets the theme size name.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The theme size name.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public ReadOnly Property SizeName() As String
                <DebuggerStepThrough>
                Get
                    Return Me.sizeNameB
                End Get
            End Property
            ''' <summary>
            ''' ( Backing Field )
            ''' The theme size name.
            ''' </summary>
            Private ReadOnly sizeNameB As String

#End Region

#Region " Constructors "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Initializes a new instance of the <see cref="ThemeInfo"/> class.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <param name="filepath">
            ''' The theme filepath.
            ''' </param>
            ''' 
            ''' <param name="colorSchemeName">
            ''' The theme color scheme name.
            ''' </param>
            ''' 
            ''' <param name="sizeName">The theme size name.
            ''' </param>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <exception cref="ArgumentNullException">
            ''' filepath or colorSchemeName or sizeName
            ''' </exception>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
            Public Sub New(ByVal filepath As String,
                           ByVal colorSchemeName As String,
                           ByVal sizeName As String)

                If String.IsNullOrWhiteSpace(filepath) Then
                    Throw New ArgumentNullException(paramName:="filepath")

                ElseIf String.IsNullOrWhiteSpace(colorSchemeName) Then
                    Throw New ArgumentNullException(paramName:="colorSchemeName")

                ElseIf String.IsNullOrWhiteSpace(sizeName) Then
                    Throw New ArgumentNullException(paramName:="sizeName")

                Else
                    Me.filepathB = filepath
                    Me.colorSchemeNameB = colorSchemeName
                    Me.sizeNameB = sizeName

                End If

            End Sub

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Prevents a default instance of the <see cref="ThemeInfo"/> class from being created.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
            Private Sub New()
            End Sub

#End Region

        End Class

#End Region

#End Region

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="EnvironmentUtil.Theming.ThemeInfo"/> object that contains the info of the current windows theme.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A <see cref="EnvironmentUtil.Theming.ThemeInfo"/> object that contains the info of the current windows theme.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property CurrentTheme() As ThemeInfo
            <DebuggerStepThrough>
            Get
                Return EnvironmentUtil.Theming.GetCurrentThemeInfo()
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the filepath of the current desktop wallpaper.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The filepath of the current desktop wallpaper.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property CurrentWallpaper() As String
            <DebuggerStepThrough>
            Get
                Dim sb As New StringBuilder(capacity:=260)

                Dim uiAction As EnvironmentUtil.NativeMethods.SystemParametersActionFlags =
                    EnvironmentUtil.NativeMethods.SystemParametersActionFlags.GetDesktopWallpaper

                If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(uiAction, CUInt(sb.Capacity), sb, Nothing) Then
                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                Else
                    Return sb.ToString
                End If
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value indicating whether Aero feature is enabled on the current operating system.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if Aero feature is enabled on the current operating system.
        ''' <see langword="False"/> if Aero feature is not enabled or else is not supported by the current operating system (like Windows XP or previous versions).
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property AeroEnabled() As Boolean
            <DebuggerStepThrough>
            Get
                If Environment.OSVersion.Version.Major < 6 Then
                    Return False ' Windows version is below Windows Vista so not Aero disponible.

                Else
                    Dim isEnabled As Boolean
                    EnvironmentUtil.NativeMethods.DwmIsCompositionEnabled(isEnabled)
                    Return isEnabled

                End If
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value indicating whether Aero feature is supported by the current operating system.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if Aero feature is supported by the current operating system.
        ''' <see langword="False"/> if Aero feature is not supported by the current operating system.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property AeroSupported() As Boolean
            <DebuggerStepThrough>
            Get
                Return (Environment.OSVersion.Version.Major > 5) ' Windows version is above Windows XP.
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value that determines wheter jpeg files are supported as wallpaper in the current operating system. 
        ''' The jpeg wallpapers are not supported before Windows Vista.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if jpeg files are supported as wallpaper in the current operating system, otherwise, <see langword="False"/>
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property WallpaperAsJpegIsSupported() As Boolean
            <DebuggerStepThrough>
            Get
                Return Environment.OSVersion.Version >= New Version(6, 0)
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value that determines whether the <see cref="WallpaperStyle.Fit"/> and <see cref="WallpaperStyle.Fill"/> are 
        ''' supported in the current operating system. 
        ''' 
        ''' The <see cref="WallpaperStyle.Fit"/> and <see cref="WallpaperStyle.Fill"/> wallpaper styles are not supported before Windows 7.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' <see langword="True"/> if <see cref="WallpaperStyle.Fit"/> and <see cref="WallpaperStyle.Fill"/> are supported in the current operating system, 
        ''' otherwise, <see langword="False"/>
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Shared ReadOnly Property WallpaperStylesFitFillAreSupported() As Boolean
            <DebuggerStepThrough>
            Get
                Return Environment.OSVersion.Version >= New Version(6, 1)
            End Get
        End Property

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Describes a wallpaper style.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Enum WallpaperStyle As Integer

            ''' <summary>
            ''' If the image is smaller than the screen, this style puts a clone of the image across the screen background.
            ''' </summary>
            Tile = 0

            ''' <summary>
            ''' Centers the image on the screen.
            ''' </summary>
            Center = 1

            ''' <summary>
            ''' Shrinks or enlarges the image to fit the monitor's height and widht. 
            ''' </summary>
            Stretch = 2

            ''' <summary>
            ''' Shrinks or enlarges the image to fit the monitor's height.
            ''' </summary>
            Fit = 3

            ''' <summary>
            ''' Shrinks or enlarges the image to fit the monitor's width.
            ''' </summary>
            Fill = 4

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Specifies a cursor type.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Enum CursorType As UInteger

            ''' <summary>
            ''' Standard arrow and small hourglass.
            ''' </summary>
            AppStarting = EnvironmentUtil.NativeMethods.SystemCursorId.AppStarting

            ''' <summary>
            ''' Standard arrow.
            ''' </summary>
            Arrow = EnvironmentUtil.NativeMethods.SystemCursorId.Arrow

            ''' <summary>
            ''' Crosshair.
            ''' </summary>
            Crosshair = EnvironmentUtil.NativeMethods.SystemCursorId.Crosshair

            ''' <summary>
            ''' Hand.
            ''' </summary>
            Hand = EnvironmentUtil.NativeMethods.SystemCursorId.Hand

            ''' <summary>
            ''' Arrow and question mark.
            ''' </summary>
            Help = EnvironmentUtil.NativeMethods.SystemCursorId.Help

            ''' <summary>
            ''' I-beam.
            ''' </summary>
            IBeam = EnvironmentUtil.NativeMethods.SystemCursorId.IBeam

            ''' <summary>
            ''' Slashed circle.
            ''' </summary>
            No = EnvironmentUtil.NativeMethods.SystemCursorId.No

            ''' <summary>
            ''' Four-pointed arrow pointing north, south, east, and west.
            ''' </summary>
            SizeAll = EnvironmentUtil.NativeMethods.SystemCursorId.SizeAll

            ''' <summary>
            ''' Double-pointed arrow pointing northeast and southwest.
            ''' </summary>
            Size_NESW = EnvironmentUtil.NativeMethods.SystemCursorId.Size_NESW

            ''' <summary>
            ''' Double-pointed arrow pointing north and south.
            ''' </summary>
            Size_NS = EnvironmentUtil.NativeMethods.SystemCursorId.Size_NS

            ''' <summary>
            ''' Double-pointed arrow pointing northwest and southeast.
            ''' </summary>
            Size_NWSE = EnvironmentUtil.NativeMethods.SystemCursorId.Size_NWSE

            ''' <summary>
            ''' Double-pointed arrow pointing west and east.
            ''' </summary>
            Size_WE = EnvironmentUtil.NativeMethods.SystemCursorId.Size_WE

            ''' <summary>
            ''' Vertical arrow.
            ''' </summary>
            Up = EnvironmentUtil.NativeMethods.SystemCursorId.Up

            ''' <summary>
            ''' Hourglass.
            ''' </summary>
            Wait = EnvironmentUtil.NativeMethods.SystemCursorId.Wait

        End Enum

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="Theming"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the current desktop wallpaper.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="imagePath">
        ''' The wallpaper filepath.
        ''' </param>
        ''' 
        ''' <param name="style">
        ''' The wallpaper style to apply.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="ArgumentNullException">
        ''' imagepath
        ''' </exception>
        ''' 
        ''' <exception cref="ArgumentException">
        ''' Invalid enumeration value;style
        ''' </exception>
        ''' 
        ''' <exception cref="Exception">
        ''' The current operating system doesn't support a fitted or filled wallpaper.
        ''' </exception>
        ''' 
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub SetDesktopWallpaper(ByVal imagePath As String,
                                              ByVal style As WallpaperStyle)

            If String.IsNullOrWhiteSpace(imagePath) Then
                Throw New ArgumentNullException(paramName:="imagepath")

            Else
                ' Set the wallpaper style and tile. 
                ' Two registry values are set in the 'HKxx\Control Panel\Desktop' key.
                '
                ' TileWallpaper:
                '  0: The wallpaper picture should not be tiled .
                '  1: The wallpaper picture should be tiled .
                ' 
                ' WallpaperStyle:
                '  0:  The image is centered if 'TileWallpaper=0' or tiled if 'TileWallpaper=1'.
                '  2:  The image is stretched to fill the screen.
                '  6:  The image is resized to fit the screen while maintaining the aspect ratio. (Windows 7 and higher)
                ' 10: The image is resized and cropped to fill the screen while maintaining the aspect ratio. (Windows 7 and higher)
                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop", writable:=True)

                    Select Case style

                        Case WallpaperStyle.Tile
                            regKey.SetValue("WallpaperStyle", "0")
                            regKey.SetValue("TileWallpaper", "1")

                        Case WallpaperStyle.Center
                            regKey.SetValue("WallpaperStyle", "0")
                            regKey.SetValue("TileWallpaper", "0")

                        Case WallpaperStyle.Stretch
                            regKey.SetValue("WallpaperStyle", "2")
                            regKey.SetValue("TileWallpaper", "0")

                        Case WallpaperStyle.Fit ' (Windows 7 and higher)
                            regKey.SetValue("WallpaperStyle", "6")
                            regKey.SetValue("TileWallpaper", "0")

                        Case WallpaperStyle.Fill ' (Windows 7 and higher)
                            regKey.SetValue("WallpaperStyle", "10")
                            regKey.SetValue("TileWallpaper", "0")

                        Case Else
                            Throw New ArgumentException(message:="Invalid enumeration value", paramName:="style")

                    End Select

                End Using

                Dim imageExt As String = Path.GetExtension(imagePath)

                If (imageExt.Equals(".jpg", StringComparison.OrdinalIgnoreCase) OrElse imageExt.Equals(".jpeg", StringComparison.OrdinalIgnoreCase)) AndAlso
                    Not (EnvironmentUtil.Theming.WallpaperStylesFitFillAreSupported) Then

                    Throw New Exception(message:="The current operating system doesn't support a fitted or filled wallpaper.")

                Else
                    Dim uiAction As EnvironmentUtil.NativeMethods.SystemParametersActionFlags =
                        EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDesktopWallpaper

                    If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(uiAction, Nothing, imagePath, EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then
                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                End If

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Removes the current desktop wallpaper from screen.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Win32Exception">
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub RemoveDesktopWallpaper()

            Dim uiAction As EnvironmentUtil.NativeMethods.SystemParametersActionFlags =
                EnvironmentUtil.NativeMethods.SystemParametersActionFlags.SetDesktopWallpaper

            If Not EnvironmentUtil.NativeMethods.SystemParametersInfo(uiAction, Nothing, String.Empty, EnvironmentUtil.NativeMethods.SystemParametersWinIniFlags.None) Then
                Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the system cursor.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' EnvironmentUtil.Theming.SetSystemCursor("C:\Windows\Cursors\aero_pen.cur", EnvironmentUtil.Theming.CursorType.Arrow)
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filePath">
        ''' The cursor file path.
        ''' </param>
        ''' 
        ''' <param name="cursorType">
        ''' The cursor type.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub SetSystemCursor(ByVal filePath As String,
                                          ByVal cursorType As EnvironmentUtil.Theming.CursorType)

            If Not EnvironmentUtil.NativeMethods.SetSystemCursor(EnvironmentUtil.NativeMethods.LoadCursorFromFile(filePath), cursorType) Then
                Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Sets the system visual theme.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <example> This is a code example.
        ''' <code>
        ''' EnvironmentUtil.Theming.SetSystemVisualTheme("C:\ThemeName.msstyles", "NormalColor", "NormalSize")
        ''' </code>
        ''' </example>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="filePath">
        ''' The theme file path.
        ''' </param>
        ''' 
        ''' <param name="colorName">
        ''' The coor scheme name.
        ''' </param>
        ''' 
        ''' <param name="sizeName">
        ''' The size name.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Shared Sub SetSystemVisualTheme(ByVal filePath As String,
                                               ByVal colorName As String,
                                               ByVal sizeName As String)

            EnvironmentUtil.NativeMethods.SetSystemVisualStyle(filePath, colorName, sizeName, 65)

        End Sub

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a <see cref="EnvironmentUtil.Theming.ThemeInfo"/> object that contains the info of the current windows theme.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' A <see cref="EnvironmentUtil.Theming.ThemeInfo"/> object that contains the info of the current windows theme.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Shared Function GetCurrentThemeInfo() As EnvironmentUtil.Theming.ThemeInfo

            Dim bufferLength As Integer = 260

            Dim sbFilepath As New StringBuilder(capacity:=bufferLength)
            Dim sbColorSchemeName As New StringBuilder(capacity:=bufferLength)
            Dim sbSizeName As New StringBuilder(capacity:=bufferLength)

            EnvironmentUtil.NativeMethods.GetCurrentThemeName(sbFilepath, bufferLength,
                                                              sbColorSchemeName, bufferLength,
                                                              sbSizeName, bufferLength)

            Return New EnvironmentUtil.Theming.ThemeInfo(filepath:=sbFilepath.ToString,
                                                         colorSchemeName:=sbColorSchemeName.ToString,
                                                         sizeName:=sbSizeName.ToString)

        End Function

#End Region

    End Class

#End Region

#End Region

End Class

#End Region
