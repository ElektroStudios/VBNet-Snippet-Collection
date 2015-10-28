
' ***********************************************************************
' Author   : Elektro
' Modified : 26-October-2015
' ***********************************************************************
' <copyright file="System Notifier.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Methods "

' SystemNotifier.NotifyDirectoryAttributesChanged(String)
' SystemNotifier.NotifyDirectoryCreated(String)
' SystemNotifier.NotifyDirectoryDeleted(String)
' SystemNotifier.NotifyDirectoryRenamed(String, String)
' SystemNotifier.NotifyDriveAdded(String, Boolean)
' SystemNotifier.NotifyDriveRemoved(String)
' SystemNotifier.NotifyFileAssociationChanged()
' SystemNotifier.NotifyFileAttributesChanged(String)
' SystemNotifier.NotifyFileCreated(String)
' SystemNotifier.NotifyFileDeleted(String)
' SystemNotifier.NotifyFileRenamed(String, String)
' SystemNotifier.NotifyFreespaceChanged(String)
' SystemNotifier.NotifyMediaInserted(String)
' SystemNotifier.NotifyMediaRemoved(String)
' SystemNotifier.NotifyNetworkFolderShared(String)
' SystemNotifier.NotifyNetworkFolderUnshared(String)
' SystemNotifier.NotifyUpdateDirectory(String)
' SystemNotifier.NotifyUpdateImage()

' SystemNotifier.ReloadSystemCursors()
' SystemNotifier.ReloadSystemIcons()

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
Imports System.ComponentModel
Imports System.Linq
Imports System.Runtime.InteropServices

#End Region

#Region " System Notifier "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Notifies the system about changes on the environment to update quickly.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module SystemNotifier

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
        ''' Sends the specified message to one or more windows.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="hWnd">
        ''' A handle to the window whose window procedure will receive the message.
        ''' 
        ''' If this parameter is <see cref="SystemNotifier.NativeMethods.WindowsMessages.HWNDBROADCAST"/>, 
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
        ''' <see cref="SystemNotifier.NativeMethods.SendMessageTimeout"/> does not provide information about 
        ''' individual windows timing out if <see cref="SystemNotifier.NativeMethods.WindowsMessages.HWNDBROADCAST"/> is used.
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
        ''' and if so, whether the <see cref="SystemNotifier.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' <see cref="SystemNotifier.NativeMethods.WindowsMessages.WM_SETTINGCHANGE"/> message.
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
                          ByVal wEventId As SystemNotifier.NativeMethods.SHChangeNotifyEventID,
                          ByVal uFlags As SystemNotifier.NativeMethods.SHChangeNotifyFlags,
                          ByVal dwItem1 As String,
                          ByVal dwItem2 As String)
        End Sub

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="fuFlags"/> parameter of <see cref="SystemNotifier.NativeMethods.SendMessageTimeout"/> function.
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
        ''' Flags for <paramref name="wEventId"/> parameter of <see cref="SystemNotifier.NativeMethods.SHChangeNotify"/> method.
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
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that was created.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DirectoryCreated = &H8UI

            ''' <summary>
            ''' A folder has been removed.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that was removed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DirectoryDeleted = &H10UI

            ''' <summary>
            ''' The name of a folder has changed.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the previous pointer to an item identifier list (PIDL) or name of the folder.
            ''' <paramref name="dwItem2"/> contains the new PIDL or name of the folder.
            ''' </summary>
            DirectoryRenamed = &H20000UI

            ''' <summary>
            ''' A non-folder item has been created.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the item that was created.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            ItemCreated = &H2UI

            ''' <summary>
            ''' A nonfolder item has been deleted.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the item that was deleted.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            ItemDeleted = &H4UI

            ''' <summary>
            ''' The name of a nonfolder item has changed.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the previous PIDL or name of the item.
            ''' <paramref name="dwItem2"/> contains the new PIDL or name of the item.
            ''' </summary>
            ItemRenamed = &H1UI

            ''' <summary>
            ''' A drive has been added.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive that was added.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DriveAdded = &H100UI

            ''' <summary>
            ''' A drive has been added and the Shell should create a new window for the drive.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive that was added.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DriveAddedShell = &H10000UI

            ''' <summary>
            ''' A drive has been removed. 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive that was removed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            DriveRemoved = &H80UI

            ''' <summary>
            ''' Storage media has been inserted into a drive.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive that contains the new media.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            MediaInserted = &H20UI

            ''' <summary>
            ''' Storage media has been removed from a drive.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive from which the media was removed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            MediaRemoved = &H40UI

            ''' <summary>
            ''' A folder on the local computer is being shared via the network.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that is being shared.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            NetShared = &H200UI

            ''' <summary>
            ''' A folder on the local computer is no longer being shared via the network.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that is no longer being shared.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            NetUnshared = &H400UI

            ''' <summary>
            ''' The computer has disconnected from a server.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the server from which the computer was disconnected.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            ServerDisconnected = &H4000UI

            ''' <summary>
            ''' The attributes of an item or folder have changed.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the item or folder that has changed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            ItemAttributesChanged = &H800UI

            ''' <summary>
            ''' A file type association has changed. 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> must be specified in the <paramref name="uFlags"/> parameter.
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are not used and must be set as <see cref="IntPtr.Zero"/>.
            ''' </summary>
            FileAssocChanged = &H8000000UI

            ''' <summary>
            ''' The amount of free space on a drive has changed.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the root of the drive on which the free space changed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' </summary>
            FreespaceChanged = &H40000UI

            ''' <summary>
            ''' The contents of an existing folder have changed but the folder still exists and has not been renamed.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIDList"/> or 
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA"/> must be specified in <paramref name="uFlags"/>.
            ''' <paramref name="dwItem1"/> contains the folder that has changed.
            ''' <paramref name="dwItem2"/> is not used and should be <see cref="IntPtr.Zero"/>.
            ''' If a folder has been created, deleted or renamed use <see cref="SystemNotifier.NativeMethods.SHChangeNotifyEventID.DirectoryCreated"/>, 
            ''' or <see cref="SystemNotifier.NativeMethods.SHChangeNotifyEventID.DirectoryRenamed"/> respectively instead.
            ''' </summary>
            UpdateDirectory = &H1000UI

            ''' <summary>
            ''' An image in the system image list has changed.
            ''' <see cref="SystemNotifier.NativeMethods.SHChangeNotifyFlags.DWORD"/> must be specified in <paramref name="uFlags"/>.
            ''' </summary>
            UpdateImage = &H8000UI

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="uFlags"/> parameter of <see cref="SystemNotifier.NativeMethods.SHChangeNotify"/> method.
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
            HwndBroadcast = &HFFFF&

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' A message that is sent to all top-level windows when 
            ''' the SystemParametersInfo function changes a system-wide setting or when policy settings have changed.
            ''' 
            ''' Applications should send <see cref="SystemNotifier.NativeMethods.WindowsMessages.WMSETTINGCHANGE"/> to all top-level windows when 
            ''' they make changes to system parameters
            ''' (This message cannot be sent directly to a single window.)
            ''' 
            ''' To send the <see cref="SystemNotifier.NativeMethods.WindowsMessages.WMSETTINGCHANGE"/> message to all top-level windows, 
            ''' use the <see cref="SystemNotifier.NativeMethods.SendMessageTimeout"/> function with the <paramref name="hwnd"/> parameter set to 
            ''' <see cref="SystemNotifier.NativeMethods.WindowsMessages.HWNDBROADCAST"/>.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <remarks>
            ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms725497%28v=vs.85%29.aspx"/>
            ''' </remarks>
            ''' ----------------------------------------------------------------------------------------------------
            WMSettingchange = &H1A

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="uiAction"/> parameter of <see cref="SystemNotifier.NativeMethods.SystemParametersInfo"/> function.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947(v=vs.85).aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Enum SystemParametersActionFlags As UInteger

            ' *****************************************************************************
            '                            WARNING!, NEED TO KNOW...
            '
            '  THIS ENUMERATION IS PARTIALLY DEFINED JUST FOR THE PURPOSES OF THIS PROJECT
            ' *****************************************************************************

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

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="fWinIni"/> parameter of <see cref="SystemNotifier.NativeMethods.SystemParametersInfo"/> function.
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
            ''' Broadcasts the <see cref="SystemNotifier.NativeMethods.WindowsMessages.WMSETTINGCHANGE"/> message after updating the user profile.
            ''' </summary>
            SendChange = &H2

            ''' <summary>
            ''' Same as <see cref="SystemNotifier.NativeMethods.SystemParametersWinIniFlags.SendChange"/>.
            ''' </summary>
            SendWinIniChange = &H3

        End Enum

#End Region

    End Class

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
    Public Sub NotifyRegistryChange(ByVal keyName As String)

        SystemNotifier.NativeMethods.SendMessageTimeout(New IntPtr(SystemNotifier.NativeMethods.WindowsMessages.HwndBroadcast),
                                                        CInt(SystemNotifier.NativeMethods.WindowsMessages.WMSettingchange),
                                                        New IntPtr(0),
                                                        keyName,
                                                        SystemNotifier.NativeMethods.SendMessageTimeoutFlags.AbortIfHung,
                                                        1,
                                                        IntPtr.Zero)

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
    Public Sub NotifyDirectoryCreated(ByVal directoryPath As String)

        If String.IsNullOrWhiteSpace(directoryPath) Then
            Throw New ArgumentNullException(paramName:="directoryPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.DirectoryCreated,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyDirectoryDeleted(ByVal directoryPath As String)

        If String.IsNullOrWhiteSpace(directoryPath) Then
            Throw New ArgumentNullException(paramName:="directoryPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.DirectoryDeleted,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyDirectoryRenamed(ByVal oldDirectoryPath As String, ByVal newDirectoryPath As String)

        If String.IsNullOrWhiteSpace(oldDirectoryPath) Then
            Throw New ArgumentNullException(paramName:="oldDirectoryPath")

        ElseIf String.IsNullOrWhiteSpace(newDirectoryPath) Then
            Throw New ArgumentNullException(paramName:="newDirectoryPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.DirectoryRenamed,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyFileCreated(ByVal filePath As String)

        If String.IsNullOrWhiteSpace(filePath) Then
            Throw New ArgumentNullException(paramName:="filePath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.ItemCreated,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyFileDeleted(ByVal filePath As String)

        If String.IsNullOrWhiteSpace(filePath) Then
            Throw New ArgumentNullException(paramName:="filePath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.ItemDeleted,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyFileRenamed(ByVal oldFilePath As String, ByVal newFilePath As String)

        If String.IsNullOrWhiteSpace(oldFilePath) Then
            Throw New ArgumentNullException(paramName:="oldFilePath")

        ElseIf String.IsNullOrWhiteSpace(newFilePath) Then
            Throw New ArgumentNullException(paramName:="newFilePath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.ItemRenamed,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyDriveAdded(ByVal driveRootPath As String,
                                Optional ByVal createShellWindow As Boolean = False)

        If String.IsNullOrWhiteSpace(driveRootPath) Then
            Throw New ArgumentNullException(paramName:="driveRootPath")

        Else
            Select Case createShellWindow

                Case True
                    SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.DriveAddedShell,
                                                                SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
                                                                dwItem1:=driveRootPath, dwItem2:=Nothing)
                Case Else
                    SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.DriveAdded,
                                                                SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyDriveRemoved(ByVal driveRootPath As String)

        If String.IsNullOrWhiteSpace(driveRootPath) Then
            Throw New ArgumentNullException(paramName:="driveRootPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.DriveRemoved,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyMediaInserted(ByVal driveRootPath As String)

        If String.IsNullOrWhiteSpace(driveRootPath) Then
            Throw New ArgumentNullException(paramName:="driveRootPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.MediaInserted,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyMediaRemoved(ByVal driveRootPath As String)

        If String.IsNullOrWhiteSpace(driveRootPath) Then
            Throw New ArgumentNullException(paramName:="driveRootPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.MediaRemoved,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyNetworkFolderShared(ByVal directoryPath As String)

        If String.IsNullOrWhiteSpace(directoryPath) Then
            Throw New ArgumentNullException(paramName:="directoryPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.NetShared,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyNetworkFolderUnshared(ByVal directoryPath As String)

        If String.IsNullOrWhiteSpace(directoryPath) Then
            Throw New ArgumentNullException(paramName:="directoryPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.NetUnshared,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyFileAttributesChanged(ByVal filePath As String)

        If String.IsNullOrWhiteSpace(filePath) Then
            Throw New ArgumentNullException(paramName:="filePath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.ItemAttributesChanged,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
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
    Public Sub NotifyDirectoryAttributesChanged(ByVal directoryPath As String)

        If String.IsNullOrWhiteSpace(directoryPath) Then
            Throw New ArgumentNullException(paramName:="directoryPath")

        Else
            SystemNotifier.NotifyFileAttributesChanged(directoryPath)

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
    Public Sub NotifyUpdateDirectory(ByVal directoryPath As String)

        If String.IsNullOrWhiteSpace(directoryPath) Then
            Throw New ArgumentNullException(paramName:="directoryPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.UpdateDirectory,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
                                                        dwItem1:=directoryPath, dwItem2:=Nothing)
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies the system that a file type association has changed.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub NotifyFileAssociationChanged()

        SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.FileAssocChanged,
                                                    SystemNotifier.NativeMethods.SHChangeNotifyFlags.ItemIdList,
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
    Public Sub NotifyFreespaceChanged(ByVal driveRootPath As String)

        If String.IsNullOrWhiteSpace(driveRootPath) Then
            Throw New ArgumentNullException(paramName:="driveRootPath")

        Else
            SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.FreespaceChanged,
                                                        SystemNotifier.NativeMethods.SHChangeNotifyFlags.PathA,
                                                        dwItem1:=driveRootPath, dwItem2:=Nothing)
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies the system that an image in the system image list has changed.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub NotifyUpdateImage()

        SystemNotifier.NativeMethods.SHChangeNotify(SystemNotifier.NativeMethods.SHChangeNotifyEventID.UpdateImage,
                                                    SystemNotifier.NativeMethods.SHChangeNotifyFlags.Dword,
                                                    dwItem1:=Nothing, dwItem2:=Nothing)

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
    Public Sub ReloadSystemCursors()

        If SystemNotifier.NativeMethods.SystemParametersInfo(
            SystemNotifier.NativeMethods.SystemParametersActionFlags.Setcursors, 0UI, 0UI,
            SystemNotifier.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
    Public Sub ReloadSystemIcons()

        If SystemNotifier.NativeMethods.SystemParametersInfo(
            SystemNotifier.NativeMethods.SystemParametersActionFlags.Seticons, 0UI, 0UI,
            SystemNotifier.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

            Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
        End If

    End Sub

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
