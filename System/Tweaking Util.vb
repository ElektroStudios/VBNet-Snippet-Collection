
' ***********************************************************************
' Author   : Elektro
' Modified : 28-October-2015
' ***********************************************************************
' <copyright file="Tweaking Util.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Child Classes "

' TweakingUtil.SystemParameters

#End Region

#Region " Properties "

' TweakingUtil.SystemParameters.ActiveWindowTrackingEnabled As Boolean
' TweakingUtil.SystemParameters.ActiveWindowTrackingTimeout As UShort
' TweakingUtil.SystemParameters.BeepEnabled As Boolean
' TweakingUtil.SystemParameters.BlockSendInputResetsEnabled As Boolean
' TweakingUtil.SystemParameters.BorderMultiplierFactor As Integer
' TweakingUtil.SystemParameters.CaretWidth As Integer
' TweakingUtil.SystemParameters.CleartypeEnabled As Boolean
' TweakingUtil.SystemParameters.ClientAreaAnimationEnabled As Boolean
' TweakingUtil.SystemParameters.ComboBoxAnimationEnabled As Boolean
' TweakingUtil.SystemParameters.CursorShadowEnabled As Boolean
' TweakingUtil.SystemParameters.DoubleClickSize As Size
' TweakingUtil.SystemParameters.DoubleClickTime As Integer
' TweakingUtil.SystemParameters.DragFullWindowsEnabled As Boolean
' TweakingUtil.SystemParameters.DragSize As Size
' TweakingUtil.SystemParameters.DropShadowEnabled As Boolean
' TweakingUtil.SystemParameters.FlatMenuEnabled As Boolean
' TweakingUtil.SystemParameters.FocusBorderSize As Size
' TweakingUtil.SystemParameters.FontSmoothingContrast As Integer
' TweakingUtil.SystemParameters.FontSmoothingEnabled As Boolean
' TweakingUtil.SystemParameters.ForegroundFlashCount As UShort
' TweakingUtil.SystemParameters.ForegroundLockTimeout As UShort
' TweakingUtil.SystemParameters.HotTrackingEnabled As Boolean
' TweakingUtil.SystemParameters.HungAppTimeout As Integer
' TweakingUtil.SystemParameters.IconSpacing As Size
' TweakingUtil.SystemParameters.IconTitleWrappingEnabled As Boolean
' TweakingUtil.SystemParameters.KeyboardDelay As Integer
' TweakingUtil.SystemParameters.KeyboardSpeed As Integer
' TweakingUtil.SystemParameters.ListBoxSmoothScrollingEnabled As Boolean
' TweakingUtil.SystemParameters.MenuAccessKeysUnderlined As Boolean
' TweakingUtil.SystemParameters.MenuAnimationEnabled As Boolean
' TweakingUtil.SystemParameters.MenuFadeEnabled As Boolean
' TweakingUtil.SystemParameters.MenuShowDelay As Integer
' TweakingUtil.SystemParameters.MessageDuration As Long
' TweakingUtil.SystemParameters.MouseButtonsSwapEnabled As Boolean
' TweakingUtil.SystemParameters.MouseClickLockEnabled As Boolean
' TweakingUtil.SystemParameters.MouseClickLockTime As Integer
' TweakingUtil.SystemParameters.MouseHoverSize As Size
' TweakingUtil.SystemParameters.MouseHoverTime As Integer
' TweakingUtil.SystemParameters.MouseSonarEnabled As Boolean
' TweakingUtil.SystemParameters.MouseSpeed As Integer
' TweakingUtil.SystemParameters.MouseTrailAmount As Integer
' TweakingUtil.SystemParameters.MouseVanishEnabled As Boolean
' TweakingUtil.SystemParameters.MouseWheelScrollLines As Integer
' TweakingUtil.SystemParameters.OverlappedContentEnabled As Boolean
' TweakingUtil.SystemParameters.PopupMenuAlignment As LeftRightAlignment
' TweakingUtil.SystemParameters.ScreensaverEnabled As Boolean
' TweakingUtil.SystemParameters.ScreensaverPath As String
' TweakingUtil.SystemParameters.ScreensaverTimeout As Integer
' TweakingUtil.SystemParameters.ScreensaveSecureEnabled As Boolean
' TweakingUtil.SystemParameters.SelectionFadeEnabled As Boolean
' TweakingUtil.SystemParameters.SnapToDefaultEnabled As Boolean
' TweakingUtil.SystemParameters.SystemDateTime As Date
' TweakingUtil.SystemParameters.SystemLanguageBarEnabled As Boolean
' TweakingUtil.SystemParameters.TitleBarGradientEnabled As Boolean
' TweakingUtil.SystemParameters.ToolTipAnimationEnabled As Boolean
' TweakingUtil.SystemParameters.UIEffectsEnabled As Boolean
' TweakingUtil.SystemParameters.WaitToKillAppTimeout As Integer
' TweakingUtil.SystemParameters.WaitToKillServiceTimeout As Integer
' TweakingUtil.SystemParameters.WheelscrollChars As Integer

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
Imports System.Linq
Imports Microsoft.Win32

#End Region

#Region " Tweaking Util "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains related Windows tweaking utilities.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module TweakingUtil

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
        ''' and if so, whether the WM_SETTINGCHANGE message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' WM_SETTINGCHANGE message.
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
        ''' and if so, whether the WM_SETTINGCHANGE message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' WM_SETTINGCHANGE message.
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
        ''' and if so, whether the WM_SETTINGCHANGE message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' WM_SETTINGCHANGE message.
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
        ''' and if so, whether the WM_SETTINGCHANGE message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' WM_SETTINGCHANGE message.
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
        ''' and if so, whether the WM_SETTINGCHANGE message is to be broadcast to 
        ''' all top-level windows to notify them of the change.
        ''' 
        ''' This parameter can be '0' if you do not want to update the user profile or broadcast the 
        ''' WM_SETTINGCHANGE message.
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

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="uiAction"/> parameter of <see cref="TweakingUtil.NativeMethods.SystemParametersInfo"/> function.
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

            ' ''' <summary>
            ' ''' Retrieves the border multiplier factor that determines the width of a window's sizing border.
            ' ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives this value.
            ' ''' </summary>
            'Getborder = &H5

            ''' <summary>
            ''' Sets the border multiplier factor that determines the width of a window's sizing border.
            ''' The <paramref name="uiParam"/> parameter specifies the new value.
            ''' </summary>
            SetBorder = &H6

            ' ''' <summary>
            ' ''' Retrieves the keyboard repeat-speed setting, which is a value in the range 
            ' ''' from 0 (approximately 2.5 repetitions per second) through 31 (approximately 30 repetitions per second). 
            ' ''' The actual repeat rates are hardware-dependent and may vary from a linear scale by as much as 20%. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Integer"/> variable that receives the setting
            ' ''' </summary>
            'GetKeyboardSpeed = &HA

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

            ' ''' <summary>
            ' ''' Sets the desktop wallpaper. 
            ' ''' The value of the <paramref name="pvParam"/> parameter determines the new wallpaper. 
            ' ''' To specify a wallpaper bitmap, set <paramref name="pvParam"/> to point to a null-terminated string containing the name of a bitmap file. 
            ' ''' Setting <paramref name="pvParam"/> to "" removes the wallpaper.
            ' ''' Setting <paramref name="pvParam"/> to null reverts to the default wallpaper.
            ' ''' </summary>
            'SetDesktopWallpaper = &H14

            ' ''' <summary>
            ' ''' Retrieves the full path of the bitmap file for the desktop wallpaper.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="StringBuilder"/> that receives a null-terminated path string.
            ' ''' Set the <paramref name="uiParam"/> parameter to the size, in characters, of the <paramref name="pvParam"/> buffer. 
            ' ''' The returned string will not exceed <see cref="StringBuilder.MaxCapacity"/> characters. 
            ' ''' If there is no desktop wallpaper, the returned string is empty.
            ' ''' </summary>
            'GetDesktopWallpaper = &H73

            ' ''' <summary>
            ' ''' Retrieves the keyboard repeat-delay setting, 
            ' ''' which is a value in the range from 0 (approximately 250 ms delay) through 3 (approximately 1 second delay). 
            ' ''' The actual delay associated with each value may vary depending on the hardware. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives the setting.
            ' ''' </summary>
            'GetKeyboardDelay = &H16

            ''' <summary>
            ''' Sets the keyboard repeat-delay setting. 
            ''' The <paramref name="uiParam"/> parameter must specify 0, 1, 2, or 3, where zero sets the shortest delay
            ''' (approximately 250 ms) and 3 sets the longest delay (approximately 1 second).
            ''' The actual delay associated with each value may vary depending on the hardware.
            ''' </summary>
            SetKeyboardDelay = &H17

            ' ''' <summary>
            ' ''' Determines whether icon-title wrapping is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, 
            ' ''' or <see langword="False"/> otherwise.
            ' ''' </summary>
            'GetIconTitleWrap = &H19

            ''' <summary>
            ''' Turns icon-title wrapping on or off. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on, or <see langword="False"/> for off.
            ''' </summary>
            SetIconTitleWrap = &H1A

            ' ''' <summary>
            ' ''' Determines whether pop-up menus are left-aligned or right-aligned, relative to the corresponding menu-bar item.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if left-aligned, 
            ' ''' or <see langword="False"/> otherwise.
            ' ''' </summary>
            'GetMenuDropAlignment = &H1B

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

            ' ''' <summary>
            ' ''' Determines whether dragging of full windows is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, or <see langword="False"/> otherwise.
            ' ''' </summary>
            'GetDragFullWindows = &H26

            ''' <summary>
            ''' Sets dragging of full windows either on or off. 
            ''' The <paramref name="uiParam"/> parameter specifies <see langword="True"/> for on, or <see langword="False"/> for off.
            ''' </summary>
            SetDragFullWindows = &H25

            ' ''' <summary>
            ' ''' Determines whether the font smoothing feature is enabled. 
            ' ''' This feature uses font antialiasing to make font curves appear smoother by painting pixels at different gray levels.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the feature is enabled,
            ' '''  or <see langword="False"/> if it is not.
            ' ''' Windows 95:  This flag is supported only if Windows Plus! is installed. See GETWINDOWSEXTENSION.
            ' ''' </summary>
            'GetFontSmoothing = &H4A

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

            ' ''' <summary>
            ' ''' Reloads the system cursors. 
            ' ''' Set the <paramref name="uiParam"/> parameter to zero and the <paramref name="pvParam"/> parameter to null.
            ' ''' </summary>
            'Setcursors = &H57

            ' ''' <summary>
            ' ''' Reloads the system icons. 
            ' ''' Set the <paramref name="uiParam"/> parameter to zero and the <paramref name="pvParam"/> parameter to null.
            ' ''' </summary>
            'Seticons = &H58

            ' ''' <summary>
            ' ''' Retrieves the input locale identifier for the system default input language. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an HKL variable that receives this value. 
            ' ''' For more information, see Languages, Locales, and Keyboard Layouts on MSDN.
            ' ''' </summary>
            'GetDefaultInputLang = &H59

            ' ''' <summary>
            ' ''' Sets the default input language for the system shell and applications. 
            ' ''' The specified language must be displayable using the current system character set. 
            ' ''' The <paramref name="pvParam"/> parameter must point to an HKL variable that contains the input locale identifier for the default language. 
            ' ''' For more information, see Languages, Locales, and Keyboard Layouts on MSDN.
            ' ''' </summary>
            'SetDefaultInputLang = &H5A

            ' ''' <summary>
            ' ''' Sets the hot key set for switching between input languages. 
            ' ''' The <paramref name="uiParam"/> and <paramref name="pvParam"/> parameters are not used.
            ' ''' The value sets the shortcut keys in the keyboard property sheets by reading the registry again. 
            ' ''' The registry must be set before this flag is used. 
            ' ''' the path in the registry is \HKEY_CURRENT_USER\keyboard layout\toggle. 
            ' ''' Valid values are "1" = ALT+SHIFT, "2" = CTRL+SHIFT, and "3" = none.
            ' ''' </summary>
            'SetLangToggle = &H5B

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

            ' ''' <summary>
            ' ''' Determines whether the snap-to-default-button feature is enabled. 
            ' ''' If enabled, the mouse cursor automatically moves to the default button, such as "OK" or "Apply", of a dialog box. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if the feature is on, 
            ' ''' or <see langword="False"/> if it is off.
            ' ''' </summary>
            'GetSnapToDefButton = &H5F

            ''' <summary>
            ''' Enables or disables the snap-to-default-button feature. 
            ''' If enabled, the mouse cursor automatically moves to the default button, such as "OK" or "Apply", of a dialog box. 
            ''' Set the <paramref name="uiParam"/> parameter to <see langword="True"/> to enable the feature, or <see langword="False"/> to disable it.
            ''' Applications should use the ShowWindow function when displaying a dialog box so the dialog manager can position the mouse cursor.
            ''' </summary>
            SetSnapToDefButton = &H60

            ' ''' <summary>
            ' ''' Retrieves the width, in pixels, of the rectangle within which the mouse pointer has to stay for TrackMouseEvent
            ' ''' to generate a WM_MOUSEHOVER message. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a UINT variable that receives the width.
            ' ''' </summary>
            'GetMouseHoverWidth = &H62

            ''' <summary>
            ''' Retrieves the width, in pixels, of the rectangle within which the mouse pointer has to stay for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' The <paramref name="pvParam"/> parameter must point to a UINT variable that receives the width.
            ''' </summary>
            SetMouseHoverWidth = &H63

            ' ''' <summary>
            ' ''' Retrieves the height, in pixels, of the rectangle within which the mouse pointer has to stay for TrackMouseEvent
            ' ''' to generate a WM_MOUSEHOVER message. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a UINT variable that receives the height.
            ' ''' </summary>
            'GetMouseHoverHeight = &H64

            ''' <summary>
            ''' Sets the height, in pixels, of the rectangle within which the mouse pointer has to stay for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' Set the <paramref name="uiParam"/> parameter to the new height.
            ''' </summary>
            SetMouseHoverHeight = &H65

            ' ''' <summary>
            ' ''' Retrieves the time, in milliseconds, that the mouse pointer has to stay in the hover rectangle for TrackMouseEvent
            ' ''' to generate a WM_MOUSEHOVER message. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a UINT variable that receives the time.
            ' ''' </summary>
            'GetMouseHoverTime = &H66

            ''' <summary>
            ''' Sets the time, in milliseconds, that the mouse pointer has to stay in the hover rectangle for TrackMouseEvent
            ''' to generate a WM_MOUSEHOVER message. 
            ''' This is used only if you pass HOVER_DEFAULT in the dwHoverTime parameter in the call to TrackMouseEvent. 
            ''' Set the <paramref name="uiParam"/> parameter to the new time.
            ''' </summary>
            SetMouseHoverTime = &H67

            ' ''' <summary>
            ' ''' Retrieves the number of lines to scroll when the mouse wheel is rotated. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UInteger"/> variable that receives the number of lines. 
            ' ''' The default value is 3.
            ' ''' </summary>
            'GetWheelScrollLines = &H68

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

            ' ''' <summary>
            ' ''' Retrieves the time, in milliseconds, that the system waits before displaying a shortcut menu when the mouse cursor is over a submenu item. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UShort"/> variable that receives the time of the delay.
            ' ''' </summary>
            'GetMenuShowDelay = &H6A

            ''' <summary>
            ''' Sets <paramref name="uiParam"/> to the time, in milliseconds, that the system waits before displaying a shortcut menu when the mouse cursor is
            ''' over a submenu item.
            ''' </summary>
            SetMenuShowDelay = &H6B

            ' ''' <summary>
            ' ''' Retrieves the current mouse speed. 
            ' ''' The mouse speed determines how far the pointer will move based on the distance the mouse moves.
            ' ''' The <paramref name="pvParam"/> parameter must point to an <see cref="Integer"/> variable that receives a value which 
            ' ''' ranges between 1 (slowest) and 20 (fastest).
            ' ''' A value of 10 is the default. 
            ' ''' The value can be set by an end user using the mouse control panel application or by an application using SETMOUSESPEED.
            ' ''' </summary>
            'GetMouseSpeed = &H70

            ''' <summary>
            ''' Sets the current mouse speed. 
            ''' The <paramref name="pvParam"/> parameter is an <see cref="Integer"/> variable between 1 (slowest) and 20 (fastest). 
            ''' A value of 10 is the default.
            ''' This value is typically set using the mouse control panel application.
            ''' </summary>
            SetMouseSpeed = &H71

            ' ''' <summary>
            ' ''' Determines whether active window tracking (activating the window the mouse is on) is on or off. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for on, 
            ' ''' or <see langword="False"/> for off.
            ' ''' </summary>
            'GetActiveWindowTracking = &H1000

            ''' <summary>
            ''' Sets active window tracking (activating the window the mouse is on) either on or off. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ''' </summary>
            SetActiveWindowTracking = &H1001

            ' ''' <summary>
            ' ''' Determines whether the menu animation feature is enabled. 
            ' ''' This master switch must be on to enable menu animation effects.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if animation is enabled 
            ' ''' and <see langword="False"/> if it is disabled.
            ' ''' If animation is enabled, GETMENUFADE indicates whether menus use fade or slide animation.
            ' ''' </summary>
            'GetMenuAnimation = &H1002

            ''' <summary>
            ''' Enables or disables menu animation. 
            ''' This master switch must be on for any menu animation to occur.
            ''' The <paramref name="pvParam"/> parameter is a <see cref="Boolean"/> variable; 
            ''' set <paramref name="pvParam"/> to <see langword="True"/> to enable animation and <see langword="False"/> to disable animation.
            ''' If animation is enabled, GETMENUFADE indicates whether menus use fade or slide animation.
            ''' </summary>
            SetMenuAnimation = &H1003

            ' ''' <summary>
            ' ''' Determines whether the slide-open effect for combo boxes is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for enabled, 
            ' ''' or <see langword="False"/> for disabled.
            ' ''' </summary>
            'GetComboboxAnimation = &H1004

            ''' <summary>
            ''' Enables or disables the slide-open effect for combo boxes. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable the gradient effect, or <see langword="False"/> to disable it.
            ''' </summary>
            SetComboboxAnimation = &H1005

            ' ''' <summary>
            ' ''' Determines whether the smooth-scrolling effect for list boxes is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point toa <see cref="Boolean"/> variable that receives <see langword="True"/> for enabled, 
            ' ''' or <see langword="False"/> for disabled.
            ' ''' </summary>
            'GetListboxSmoothScrolling = &H1006

            ''' <summary>
            ''' Enables or disables the smooth-scrolling effect for list boxes. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable the smooth-scrolling effect,
            ''' or <see langword="False"/> to disable it.
            ''' </summary>
            SetListboxSmoothScrolling = &H1007

            ' ''' <summary>
            ' ''' Determines whether the gradient effect for window title bars is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for enabled, 
            ' ''' or <see langword="False"/> for disabled. 
            ' ''' For more information about the gradient effect, see the GetSysColor function.
            ' ''' </summary>
            'GetGradientCaptions = &H1008

            ''' <summary>
            ''' Enables or disables the gradient effect for window title bars. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable it, or <see langword="False"/> to disable it.
            ''' The gradient effect is possible only if the system has a color depth of more than 256 colors. For more information about
            ''' the gradient effect, see the GetSysColor function.
            ''' </summary>
            SetGradientCaptions = &H1009

            ' ''' <summary>
            ' ''' Determines whether menu access keys are always underlined. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that 
            ' ''' receives <see langword="True"/> if menu access keys are always underlined, 
            ' ''' and <see langword="False"/> if they are underlined only when the menu is activated by the keyboard.
            ' ''' </summary>
            'GetKeyboardCues = &H100A

            ''' <summary>
            ''' Sets the underlining of menu access key letters. 
            ''' The <paramref name="pvParam"/> parameter is a <see cref="Boolean"/> variable. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to always underline menu access keys, 
            ''' or <see langword="False"/> to underline menu access keys only when the menu is activated from the keyboard.
            ''' </summary>
            SetKeyboardCues = &H100B

            ' ''' <summary>
            ' ''' Determines whether hot tracking of user-interface elements, such as menu names on menu bars, is enabled.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> for enabled, 
            ' ''' or <see langword="False"/> for disabled.
            ' ''' Hot tracking means that when the cursor moves over an item, it is highlighted but not selected. 
            ' ''' You can query this value to decide whether to use hot tracking in the user interface of your application.
            ' ''' </summary>
            'GetHotTracking = &H100E

            ''' <summary>
            ''' Enables or disables hot tracking of user-interface elements such as menu names on menu bars. 
            ''' Set the <paramref name="pvParam"/> parameter to <see langword="True"/> to enable it, or <see langword="False"/> to disable it.
            ''' Hot-tracking means that when the cursor moves over an item, it is highlighted but not selected.
            ''' </summary>
            SetHotTracking = &H100F

            ' ''' <summary>
            ' ''' Determines whether menu fade animation is enabled.
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/>
            ' ''' when fade animation is enabled and <see langword="False"/> when it is disabled. 
            ' ''' If fade animation is disabled, menus use slide animation.
            ' ''' This flag is ignored unless menu animation is enabled, which you can do using the SETMENUANIMATION flag.
            ' ''' For more information, see AnimateWindow.
            ' ''' </summary>
            'GetMenuFade = &H1012

            ''' <summary>
            ''' Enables or disables menu fade animation. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable the menu fade effect or <see langword="False"/> to disable it.
            ''' If fade animation is disabled, menus use slide animation.
            ''' The menu fade effect is possible only if the system has a color depth of more than 256 colors. 
            ''' This flag is ignored unless MENUANIMATION is also set. 
            ''' For more information, see AnimateWindow.
            ''' </summary>
            SetMenuFade = &H1013

            ' ''' <summary>
            ' ''' Determines whether the selection fade effect is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled 
            ' ''' or <see langword="False"/> if disabled.
            ' ''' The selection fade effect causes the menu item selected by the user to remain on the screen briefly while fading out
            ' ''' after the menu is dismissed.
            ' ''' </summary>
            'GetSelectionFade = &H1014

            ''' <summary>
            ''' The selection fade effect causes the menu item selected by the user to remain on the screen briefly while fading out
            ''' after the menu is dismissed. 
            ''' The selection fade effect is possible only if the system has a color depth of more than 256 colors.
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable the selection fade effect or <see langword="False"/> to disable it.
            ''' </summary>
            SetSelectionFade = &H1015

            ' ''' <summary>
            ' ''' Determines whether ToolTip animation is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled 
            ' ''' or <see langword="False"/> if disabled. 
            ' ''' If ToolTip animation is enabled, GETTOOLTIPFADE indicates whether ToolTips use fade or slide animation.
            ' ''' </summary>
            'GetTooltipAnimation = &H1016

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

            ' ''' <summary>
            ' ''' Determines whether native User menus have flat menu appearance. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that returns <see langword="True"/> if the 
            ' ''' flat menu appearance is set, or <see langword="False"/> otherwise.
            ' ''' </summary>
            'GetFlatMenu = &H1022

            ''' <summary>
            ''' Enables or disables flat menu appearance for native User menus. 
            ''' Set <paramref name="pvParam"/> to <see langword="True"/> to enable flat menu appearance or <see langword="False"/> to disable it.
            ''' When enabled, the menu bar uses COLOR_MENUBAR for the menubar background, COLOR_MENU for the menu-popup background, COLOR_MENUHILIGHT
            ''' for the fill of the current menu selection, and COLOR_HILIGHT for the outline of the current menu selection.
            ''' If disabled, menus are drawn using the same metrics and colors as in Windows 2000 and earlier.
            ''' </summary>
            SetFlatMenu = &H1023

            ' ''' <summary>
            ' ''' Determines whether the drop shadow effect is enabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that returns <see langword="True"/> if enabled or 
            ' ''' <see langword="False"/> if disabled.
            ' ''' </summary>
            'GetDropShadow = &H1024

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

            ' ''' <summary>
            ' ''' Determines whether UI effects are enabled or disabled. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/>
            ' ''' if all UI effects are enabled, or <see langword="False"/> if they are disabled.
            ' ''' </summary>
            'GetUiEffects = &H103E

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

            ' ''' <summary>
            ' ''' Retrieves the caret width in edit controls, in pixels. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UShort"/> that receives this value.
            ' ''' </summary>
            'GetCaretWidth = &H2006

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

            ' ''' <summary>
            ' ''' Retrieves a contrast value that is used in ClearType™ smoothing. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="UInteger"/> that receives the information.
            ' ''' </summary>
            'GetFontSmoothingContrast = &H200C

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
            ' ''' Same as <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetKeyboardCues"/>.
            ' ''' </summary>
            ' GetMenuUnderlines = GETKEYBOARDCUES

            ' ''' <summary>
            ' ''' Same as <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.SetKeyboardCues"/>.
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
            ' ''' Does not work for Windows 7: <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947(v=vs.85).aspx"/>
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
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetDockMoving = &H90

            ' ''' <summary>
            ' ''' Sets whether a window is docked when it is moved to the top, left, or right docking targets on a monitor or monitor array. 
            ' ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetDockMoving = &H91

            ' ''' <summary>
            ' ''' Determines whether a maximized window is restored when its caption bar is dragged. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, or <see langword="False"/> otherwise.
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
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
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetMouseDockThreshold = &H7E

            ' ''' <summary>
            ' ''' Sets the threshold in pixels where docking behavior is triggered by using a mouse to drag a window to the edge of a monitor or monitor array. 
            ' ''' The default threshold is 1. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that contains the new value
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetMouseDockThreshold = &H7F

            ' ''' <summary>
            ' ''' Retrieves the threshold in pixels where undocking behavior is triggered by using a mouse to drag a window from the edge of a monitor or 
            ' ''' a monitor array toward the center. 
            ' ''' The default threshold is 20.
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetMouseDragoutThreshold = &H84

            ' ''' <summary>
            ' ''' Sets the threshold in pixels where undocking behavior is triggered by using a mouse to drag a window from the edge of a monitor or 
            ' ''' monitor array to its center. 
            ' ''' The default threshold is 20. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that contains the new value.
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetMouseDragoutThreshold = &H85

            ' ''' <summary>
            ' ''' Retrieves the threshold in pixels from the top of a monitor or a monitor array where a vertically maximized window is restored when dragged with the mouse. 
            ' ''' The default threshold is 50.
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetMouseSideMoveThreshold = &H88

            ' ''' <summary>
            ' ''' Sets the threshold in pixels from the top of the monitor where a vertically maximized window is restored when dragged with the mouse. 
            ' ''' The default threshold is 50. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a DWORD variable that contains the new value
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' SetMouseSideMoveThreshold = &H89

            ' ''' <summary>
            ' ''' Determines whether a window is vertically maximized when it is sized to the top or bottom of a monitor or monitor array. 
            ' ''' The <paramref name="pvParam"/> parameter must point to a <see cref="Boolean"/> variable that receives <see langword="True"/> if enabled, or <see langword="False"/> otherwise.
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
            ' ''' </summary>
            ' GetSnapSizing = &H8E

            ' ''' <summary>
            ' ''' Sets whether a window is vertically maximized when it is sized to the top or bottom of the monitor. 
            ' ''' Set <paramref name="pvParam"/> to <see langword="True"/> for on or <see langword="False"/> for off.
            ' ''' Use <see cref="TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWinArranging"/>  to determine whether this behavior is enabled.
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
        ''' Flags for <paramref name="fWinIni"/> parameter of <see cref="TweakingUtil.NativeMethods.SystemParametersInfo"/> function.
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
            ''' Broadcasts the WM_SETTINGCHANGE message after updating the user profile.
            ''' </summary>
            SendChange = &H2

            ''' <summary>
            ''' Same as <see cref="TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendChange"/>.
            ''' </summary>
            SendWinIniChange = &H3

        End Enum

#End Region

    End Class

#End Region

#Region " Child Classes "

#Region " System Parameters "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains related system-parameter utilities.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class SystemParameters

#Region " Properties "

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                    TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWaitToKillServiceTimeout, 0UI, result,
                    TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 2147483647.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetWaitToKillServiceTimeout, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWaitToKillTimeout, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    '   Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 2147483647.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetWaitToKillTimeout, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetHungAppTimeout, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 2147483647.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetHungAppTimeout, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetScreensaveSecure, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetScreensaveSecure, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetWheelscrollChars, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > Integer.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 2147483647.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetWheelscrollChars, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetSystemlanguageBar, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetSystemlanguageBar, 0UI, int32value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetCleartype, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetCleartype, 0UI, int32value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetClientAreaAnimation, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetClientAreaAnimation, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetDisableOverlappedContent, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetDisableOverlappedContent, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetFontSmoothing, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetDragFullWindows, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMousebuttonSwap, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetUiEffects, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetBlockSendInputResets, 0I, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetBlockSendInputResets, Not value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetFlatMenu, 0I, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetMouseVanish, 0I, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseVanish, 0I, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetMouseClickLock, 0I, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseClickLock, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetMouseSonar, 0I, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseSonar, 0I, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetCursorShadow, 0I, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetCursorShadow, 0I, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetDropShadow, 0I, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetSelectionFade, 0I, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMenuFade, 0I, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetHotTracking, 0I, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetKeyboardCues, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetGradientCaptions, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetListboxSmoothScrolling, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetTooltipAnimation, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetComboboxAnimation, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMenuAnimation, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetActiveWindowTracking, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetSnapToDefButton, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetMouseTrails, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetScreensaveActive, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetIconTitleWrap, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetBeep, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If

                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Boolean)
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetBeep, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.UpdateIniFile) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetForegroundFlashCount, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As UShort)
                If (value < 0) OrElse (value > UShort.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 65535.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetForegroundFlashCount, 0I, CUInt(value),
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetActiveWndTrkTimeout, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As UShort)
                If (value < 0) OrElse (value > UShort.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 65535.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetActiveWndTrkTimeout, 0I, CUInt(value),
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetForegroundLockTimeout, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return result
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As UShort)
                If (value < 0) OrElse (value > UShort.MaxValue) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 65535.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetForegroundLockTimeout, 0I, CUInt(value),
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetBorder, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(TweakingUtil.NativeMethods.SystemParametersActionFlags.SetFontSmoothingContrast,
                                                                              0UI,
                                                                              value,
                                                                              TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange Or
                                                                              TweakingUtil.NativeMethods.SystemParametersWinIniFlags.UpdateIniFile) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetMouseClickLockTime, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) Then
                    Throw New ArgumentException(message:="Value should be grater than -1.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseClickLockTime, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetCaretWidth, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseSpeed, 0UI, value,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMenuShowDelay, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetWheelScrollLines, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseHoverTime, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetMouseTrails, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 0) OrElse (value > 16) Then
                    Throw New ArgumentException(message:="Value should be between 0 and 16.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseTrails, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetDoubleclickTime, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetKeyboardDelay, CUInt(value), 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetScreensaveTimeout, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CInt(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As Integer)
                If (value < 1) OrElse (value > 599940) Then
                    Throw New ArgumentException(message:="Value should be between 1 and 599940.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetScreensaveTimeout, value, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetKeyboardSpeed, CUInt(value), 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetMessageDuration, 0UI, result,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    ' Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If
                Return CULng(result)
            End Get

            <DebuggerStepThrough>
            Set(ByVal value As ULong)
                If (value < 5) OrElse (value > 4294967295) Then
                    Throw New ArgumentException(message:="Value should be between 5 and 4294967295.", paramName:="value")

                Else
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMessageDuration, 0UI, CLng(value),
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetFocusBorderWidth, 0UI, width,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

                    Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                End If

                Dim height As UInteger
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.GetFocusBorderHeight, 0UI, height,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.None) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetFocusBorderWidth, 0UI, value.Width,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetFocusBorderHeight, 0UI, value.Height,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseHoverWidth, value.Width, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMouseHoverHeight, value.Height, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetDragWidth, value.Width, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetDragHeight, value.Height, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetDoubleClickWidth, value.Width, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.SetDoubleClickHeight, value.Height, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.IconHorizontalSpacing, value.Width, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

                        Throw New Win32Exception([error]:=Marshal.GetLastWin32Error)
                    End If

                    If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                        TweakingUtil.NativeMethods.SystemParametersActionFlags.IconVerticalSpacing, value.Height, 0UI,
                        TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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
                If Not TweakingUtil.NativeMethods.SystemParametersInfo(
                    TweakingUtil.NativeMethods.SystemParametersActionFlags.SetMenuDropAlignment, CBool(value), 0UI,
                    TweakingUtil.NativeMethods.SystemParametersWinIniFlags.SendWinIniChange) Then

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

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="SystemParameters"/> class from being created.
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
